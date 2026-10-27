"""Локальный интерфейс на стандартной библиотеке Python 3.10+ (без cgi/Tkinter)."""
from __future__ import annotations

import errno
import json
import secrets
import shutil
import tempfile
import threading
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlsplit
from zipfile import BadZipFile

from .importer import ROOT, ImportProblem, import_files

ASSETS = Path(__file__).with_name("web_assets")
MAX_FILE = 256 * 1024 * 1024
MAX_TOTAL = 512 * 1024 * 1024
MAX_TEMPLATES = 500


class RequestProblem(Exception):
    def __init__(self, message, status=400):
        self.status = status
        super().__init__(message)


def validate_file(item):
    if not isinstance(item, dict):
        raise RequestProblem("Неверное описание файла.")
    name, size = item.get("name"), item.get("size")
    if not isinstance(name, str) or not name.lower().endswith(".xlsx"):
        raise RequestProblem("Выберите файлы в формате .xlsx.")
    if len(name) > 240 or any(c in name for c in ("/", "\\", "\x00", "\r", "\n")):
        raise RequestProblem("Недопустимое имя файла.")
    if type(size) is not int or not 0 < size <= MAX_FILE:
        raise RequestProblem("Файл должен быть непустым и не больше 256 МБ.")
    return {"name": name, "size": size}


@dataclass
class Job:
    id: str
    folder: Path
    manifest: list
    paths: list
    uploaded: set = field(default_factory=set)
    uploading: set = field(default_factory=set)
    state: str = "uploading"
    message: str = "Ожидаю файлы…"
    counts: dict | None = None
    error: str | None = None
    filename: str = ""


class AppState:
    def __init__(self, config_path=None, examples_root=None):
        self.config_path = config_path
        self.examples_root = Path(examples_root or ROOT / "examples")
        self.token = secrets.token_urlsafe(32)
        self.temp = tempfile.TemporaryDirectory(prefix="matrix-importer-")
        self.lock = threading.RLock()
        self.jobs = {}
        self.workers = []

    def create_job(self, data):
        if not isinstance(data, dict):
            raise RequestProblem("Ожидается описание матрицы и шаблонов.")
        templates = data.get("templates")
        if not isinstance(templates, list) or not 1 <= len(templates) <= MAX_TEMPLATES:
            raise RequestProblem("Выберите от 1 до 500 шаблонов.")
        manifest = [validate_file(data.get("matrix")), *[validate_file(x) for x in templates]]
        if sum(f["size"] for f in manifest) > MAX_TOTAL:
            raise RequestProblem("Общий объём файлов превышает 512 МБ.", 413)
        with self.lock:
            if any(j.state == "running" for j in self.jobs.values()):
                raise RequestProblem("Дождитесь завершения текущего импорта.", 409)
            # Сохраняем несколько последних результатов до остановки программы.
            if len(self.jobs) >= 8:
                old_id = next((key for key, j in self.jobs.items() if j.state in ("done", "error")), None)
                if old_id is None:
                    raise RequestProblem("Слишком много одновременных загрузок.", 409)
                old = self.jobs.pop(old_id)
                shutil.rmtree(old.folder)
            ident = secrets.token_hex(16)
            folder = Path(self.temp.name) / ident
            folder.mkdir()
            # Номера защищают от одинаковых имён и совпадения с именем результата.
            paths = [folder / f"{i:03d}_{item['name'][:-5].encode('utf-8')[:160].decode('utf-8', errors='ignore')}.xlsx" for i, item in enumerate(manifest)]
            job = Job(ident, folder, manifest, paths)
            self.jobs[ident] = job
            return job

    def get_job(self, ident):
        with self.lock:
            if ident not in self.jobs:
                raise RequestProblem("Импорт не найден. Возможно, программа была перезапущена.", 404)
            return self.jobs[ident]

    def describe(self, job):
        with self.lock:
            result = {"id": job.id, "state": job.state, "message": job.message}
            if job.error:
                result["error"] = job.error
            if job.counts is not None:
                result.update(counts=job.counts, download_url=f"/api/jobs/{job.id}/download", filename=job.filename)
            return result

    def launch(self, job):
        with self.lock:
            if job.state != "uploading" or job.uploading or len(job.uploaded) != len(job.paths):
                raise RequestProblem("Загрузка файлов ещё не завершена или импорт уже запущен.", 409)
            if any(j.state == "running" for j in self.jobs.values()):
                raise RequestProblem("Другой импорт уже выполняется.", 409)
            job.state = "running"
            job.message = "Начинаю импорт…"
            worker = threading.Thread(target=self._work, args=(job,), daemon=True)
            self.workers.append(worker)
            worker.start()

    def _work(self, job):
        def progress(message):
            with self.lock:
                job.message = message
        try:
            result = import_files(job.paths[0], job.paths[1:], job.folder / "result.xlsx", self.config_path, progress)
            with self.lock:
                job.counts = {"added": result.added, "updated": result.updated, "skipped": result.skipped, "total": result.total}
                job.filename = f"matrix_result_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                job.state, job.message = "done", "Импорт завершён"
        except Exception as exc:
            if isinstance(exc, BadZipFile):
                message = "Один из файлов не является обычным .xlsx: он повреждён, зашифрован или имеет другой формат."
            else:
                message = str(exc) or type(exc).__name__
            with self.lock:
                job.state, job.message, job.error = "error", "Импорт остановлен", message

    def demo_available(self):
        return all((self.examples_root / name).is_file() for name in (
            "01_matrix.xlsx", "02_template_update.xlsx", "03_template_new_group.xlsx"))

    def demo(self):
        if not self.demo_available():
            raise RequestProblem("Папка examples не перенесена. Выберите свои файлы.", 404)
        sources = [self.examples_root / name for name in ("01_matrix.xlsx", "02_template_update.xlsx", "03_template_new_group.xlsx")]
        items = [{"name": p.name, "size": p.stat().st_size} for p in sources]
        job = self.create_job({"matrix": items[0], "templates": items[1:]})
        for i, (source, target) in enumerate(zip(sources, job.paths)):
            shutil.copyfile(source, target)
            job.uploaded.add(i)
        self.launch(job)
        return job

    def close(self):
        # При Ctrl+C даём текущему импорту закончить работу с временными файлами.
        for worker in self.workers:
            worker.join()
        self.temp.cleanup()


class LocalServer(ThreadingHTTPServer):
    daemon_threads = False
    block_on_close = True

    def __init__(self, port, state):
        self.app = state
        super().__init__(("127.0.0.1", port), Handler)


class Handler(BaseHTTPRequestHandler):
    server_version = "MatrixImporter/0.2"
    protocol_version = "HTTP/1.0"

    def log_message(self, *_):
        pass

    def setup(self):
        super().setup()
        self.connection.settimeout(120)

    def _guard(self, api=False):
        expected = f"127.0.0.1:{self.server.server_port}"
        if self.headers.get("Host") != expected:
            raise RequestProblem("Откройте локальный адрес, напечатанный в терминале.", 403)
        origin = self.headers.get("Origin")
        if origin and origin != f"http://{expected}":
            raise RequestProblem("Запрос с другого сайта отклонён.", 403)
        if api and not secrets.compare_digest(self.headers.get("X-Matrix-Token", ""), self.server.app.token):
            raise RequestProblem("Обновите страницу приложения.", 403)

    def _send(self, status, content, content_type="application/json; charset=utf-8", extra=None):
        if isinstance(content, dict):
            content = json.dumps(content, ensure_ascii=False).encode("utf-8")
        elif isinstance(content, str):
            content = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self' data:; object-src 'none'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(content)

    def _length(self, limit):
        try:
            length = int(self.headers.get("Content-Length", "-1"))
        except ValueError:
            length = -1
        if length < 0 or self.headers.get("Transfer-Encoding"):
            raise RequestProblem("Не указан размер запроса.", 411)
        if length > limit:
            raise RequestProblem("Слишком большой запрос.", 413)
        return length

    def _json(self):
        try:
            return json.loads(self.rfile.read(self._length(65536)).decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as exc:
            raise RequestProblem("Не удалось прочитать запрос.") from exc

    def _parts(self):
        path = urlsplit(self.path).path
        parts = path.strip("/").split("/")
        return path, parts

    def _handle(self, method):
        try:
            self._guard(api=urlsplit(self.path).path.startswith("/api/"))
            self._route(method)
        except RequestProblem as exc:
            self._send(exc.status, {"error": str(exc)})
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as exc:
            try:
                self._send(500, {"error": f"Не удалось обработать запрос: {exc}"})
            except (BrokenPipeError, ConnectionResetError):
                pass

    def do_GET(self):
        self._handle("GET")

    def do_POST(self):
        self._handle("POST")

    def do_PUT(self):
        self._handle("PUT")

    def _route(self, method):
        path, parts = self._parts()
        app = self.server.app
        if method == "GET" and path == "/":
            page = (ASSETS / "index.html").read_text(encoding="utf-8").replace("__TOKEN__", app.token)
            return self._send(200, page, "text/html; charset=utf-8")
        assets = {"/app.js": ("app.js", "text/javascript; charset=utf-8"), "/style.css": ("style.css", "text/css; charset=utf-8")}
        if method == "GET" and path in assets:
            name, mime = assets[path]
            return self._send(200, (ASSETS / name).read_bytes(), mime)
        if method == "GET" and path == "/api/info":
            return self._send(200, {"demo": app.demo_available(), "version": "0.2"})
        if method == "POST" and path == "/api/jobs":
            return self._send(201, app.describe(app.create_job(self._json())))
        if method == "POST" and path == "/api/demo":
            self._json()
            return self._send(202, app.describe(app.demo()))
        if len(parts) >= 3 and parts[:2] == ["api", "jobs"]:
            job = app.get_job(parts[2])
            if method == "GET" and len(parts) == 3:
                return self._send(200, app.describe(job))
            if method == "PUT" and len(parts) == 5 and parts[3] == "files":
                return self._upload(job, parts[4])
            if method == "POST" and len(parts) == 4 and parts[3] == "start":
                self._json()
                app.launch(job)
                return self._send(202, app.describe(job))
            if method == "GET" and len(parts) == 4 and parts[3] == "download":
                with app.lock:
                    if job.state != "done":
                        raise RequestProblem("Результат ещё не готов.", 409)
                    data = (job.folder / "result.xlsx").read_bytes()
                    name = job.filename
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", {
                    "Content-Disposition": f"attachment; filename=\"{name}\"; filename*=UTF-8''{quote(name)}"})
        raise RequestProblem("Страница не найдена.", 404)

    def _upload(self, job, slot):
        try:
            index = int(slot)
        except ValueError as exc:
            raise RequestProblem("Неверный номер файла.") from exc
        app = self.server.app
        length = self._length(MAX_FILE)
        with app.lock:
            if job.state != "uploading" or index in job.uploading or index in job.uploaded:
                raise RequestProblem("Файл уже принят или импорт запущен.", 409)
            if not 0 <= index < len(job.paths):
                raise RequestProblem("Неверный номер файла.")
            if length != job.manifest[index]["size"]:
                raise RequestProblem("Размер файла не совпадает с заявленным.")
            job.uploading.add(index)
        path = job.paths[index]
        try:
            remaining = length
            with path.open("xb") as stream:
                while remaining:
                    chunk = self.rfile.read(min(1024 * 1024, remaining))
                    if not chunk:
                        raise RequestProblem("Загрузка прервана. Повторите импорт.")
                    stream.write(chunk)
                    remaining -= len(chunk)
            with app.lock:
                job.uploaded.add(index)
        except Exception:
            path.unlink(missing_ok=True)
            raise
        finally:
            with app.lock:
                job.uploading.discard(index)
        self._send(200, {"ok": True})


def start(config_path=None, port=8765, open_browser=True):
    app = AppState(config_path)
    try:
        try:
            server = LocalServer(port, app)
        except OSError as exc:
            if exc.errno != errno.EADDRINUSE or port == 0:
                raise
            server = LocalServer(0, app)
        url = f"http://127.0.0.1:{server.server_port}"
        print(f"MatrixImporter 0.2\nОткройте в браузере: {url}\nДля завершения нажмите Ctrl+C в этом терминале.", flush=True)
        if open_browser:
            def launch_browser():
                try:
                    webbrowser.open(url)
                except Exception:
                    pass
            threading.Thread(target=launch_browser, daemon=True).start()
        try:
            server.serve_forever(poll_interval=0.25)
        except KeyboardInterrupt:
            print("\nЗавершаю работу…", flush=True)
        finally:
            server.server_close()
    finally:
        app.close()
