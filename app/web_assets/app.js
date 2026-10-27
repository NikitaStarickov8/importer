"use strict";
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="matrix-token"]').content;
let matrix = null, templates = [], busy = false, currentResult = null;
const controls = [$("matrix-file"), $("template-files"), $("run"), $("demo")];

function sizeLabel(bytes) { return bytes < 1024 * 1024 ? `${Math.ceil(bytes/1024)} КБ` : `${(bytes/1024/1024).toFixed(1)} МБ`; }
function clearResult() { currentResult = null; $("status-panel").hidden = true; }
function refresh() {
  $("matrix-name").textContent = matrix ? matrix.name : "Файл не выбран";
  $("matrix-detail").textContent = matrix ? sizeLabel(matrix.size) : "Один файл .xlsx";
  $("templates-count").textContent = templates.length ? `Выбрано шаблонов: ${templates.length}` : "Шаблоны не выбраны";
  const list = $("file-list"); list.replaceChildren(); list.hidden = !templates.length;
  templates.forEach((file, index) => {
    const item = document.createElement("li"), label = document.createElement("span"), size = document.createElement("span"), remove = document.createElement("button");
    label.className = "file-label"; label.textContent = file.name;
    size.className = "file-size"; size.textContent = sizeLabel(file.size); label.append(size);
    remove.className = "remove"; remove.textContent = "×"; remove.type = "button"; remove.setAttribute("aria-label", `Убрать ${file.name}`); remove.disabled = busy;
    remove.addEventListener("click", () => { templates.splice(index, 1); clearResult(); refresh(); });
    item.append(label, remove); list.append(item);
  });
  controls.forEach(control => control.disabled = busy);
  $("run").disabled = busy || !matrix || !templates.length;
}
function setBusy(value) { busy = value; refresh(); }
function status(message, state = "running") {
  const panel = $("status-panel"); panel.hidden = false; panel.className = `status-panel ${state}`;
  $("status-heading").textContent = state === "error" ? "Импорт не выполнен" : state === "done" ? "Матрица готова" : "Выполняется импорт";
  $("status-badge").textContent = state === "error" ? "Ошибка" : state === "done" ? "Готово" : "В работе";
  $("status-message").textContent = message;
  $("progress").hidden = state !== "running";
  if (state !== "done") { $("counts").hidden = true; $("download-row").hidden = true; }
}
async function api(path, options = {}) {
  let response;
  try { response = await fetch(path, {...options, headers: {"X-Matrix-Token": token, ...options.headers}}); }
  catch { throw new Error("Нет связи с локальной программой. Убедитесь, что python3 run.py продолжает работать в терминале."); }
  if (!response.ok) {
    let message = `Не удалось выполнить запрос (${response.status}).`;
    try { message = (await response.json()).error || message; } catch {}
    throw new Error(message);
  }
  return response;
}
async function post(path, data = {}) { return (await api(path, {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(data)})).json(); }
function fileInfo(file) { return {name:file.name, size:file.size}; }
function checkFile(file) {
  if (!file.name.toLowerCase().endsWith(".xlsx")) throw new Error(`Файл «${file.name}»: поддерживается формат .xlsx.`);
  if (!file.size || file.size > 256*1024*1024) throw new Error(`Файл «${file.name}» должен быть непустым и не больше 256 МБ.`);
}
async function watch(id) {
  for (;;) {
    const job = await (await api(`/api/jobs/${id}`)).json();
    if (job.state === "error") throw new Error(job.error || job.message);
    if (job.state === "done") {
      currentResult = job;
      status("Все выбранные шаблоны обработаны. Скачайте обновлённую матрицу.", "done");
      for (const field of ["added", "updated", "skipped"]) $(field).textContent = job.counts[field];
      $("counts").hidden = false; $("download-row").hidden = false;
      return;
    }
    status(job.message);
    await new Promise(resolve => setTimeout(resolve, 700));
  }
}
$("matrix-file").addEventListener("change", event => {
  const selected = event.target.files[0];
  if (!selected) return;
  try { checkFile(selected); matrix = selected; clearResult(); refresh(); }
  catch (error) { event.target.value = ""; status(error.message, "error"); }
});
$("template-files").addEventListener("change", event => {
  const selected = [...event.target.files];
  try {
    selected.forEach(checkFile);
    const next = [...templates];
    for (const file of selected) if (!next.some(old => old.name === file.name && old.size === file.size && old.lastModified === file.lastModified)) next.push(file);
    if (next.length > 500) throw new Error("За один импорт можно выбрать до 500 шаблонов.");
    templates = next; clearResult(); refresh();
  } catch (error) { status(error.message, "error"); }
  event.target.value = "";
});
$("run").addEventListener("click", async () => {
  if (busy || !matrix || !templates.length) return;
  setBusy(true); currentResult = null;
  try {
    const files = [matrix, ...templates];
    if (files.reduce((sum, f) => sum + f.size, 0) > 512*1024*1024) throw new Error("Общий объём файлов превышает 512 МБ.");
    status("Подготавливаю файлы…");
    const job = await post("/api/jobs", {matrix:fileInfo(matrix), templates:templates.map(fileInfo)});
    for (let i=0; i<files.length; i++) {
      status(`Передаю локальной программе файл ${i+1} из ${files.length}: ${files[i].name}`);
      await api(`/api/jobs/${job.id}/files/${i}`, {method:"PUT", headers:{"Content-Type":"application/octet-stream"}, body:files[i]});
    }
    await post(`/api/jobs/${job.id}/start`);
    await watch(job.id);
  } catch (error) { status(error.message, "error"); }
  finally { setBusy(false); }
});
$("demo").addEventListener("click", async () => {
  if (busy) return;
  setBusy(true); currentResult = null; status("Обрабатываю искусственные примеры…");
  try { const job = await post("/api/demo"); await watch(job.id); }
  catch (error) { status(error.message, "error"); }
  finally { setBusy(false); }
});
$("download").addEventListener("click", async () => {
  if (!currentResult) return;
  const result = currentResult;
  $("download").disabled = true;
  try {
    const blob = await (await api(result.download_url)).blob();
    const url = URL.createObjectURL(blob), link = document.createElement("a");
    link.href = url; link.download = result.filename; document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (error) { $("status-message").textContent = error.message; }
  finally { $("download").disabled = false; }
});
api("/api/info").then(response => response.json()).then(info => { $("demo").hidden = !info.demo; }).catch(error => status(error.message, "error"));
refresh();
