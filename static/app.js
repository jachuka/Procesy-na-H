const $ = id => document.getElementById(id);
let file, result;
const decimal = n => n.toFixed(Number($('precision').value)).replace('.', ',');
function duration(h) {
  const seconds = Math.round(h * 3600);
  return `${Math.floor(seconds / 3600)}:${String(Math.floor(seconds % 3600 / 60)).padStart(2,'0')}:${String(seconds % 60).padStart(2,'0')}`;
}
async function copy(text) {
  try {
    if (navigator.clipboard && window.isSecureContext) await navigator.clipboard.writeText(text);
    else {
      const area = document.createElement('textarea'); area.value = text; document.body.append(area); area.select();
      const ok = document.execCommand('copy'); area.remove(); if (!ok) throw Error();
    }
    $('status').textContent = 'Skopiowano do schowka.';
  } catch { $('status').textContent = 'Kopiowanie jest niedostępne. Zaznacz liczbę w tabeli i skopiuj ją ręcznie.'; }
}
async function calculate(sheet) {
  $('results').hidden = true; result = null; $('status').textContent = 'Odczytuję wycenę…';
  $('file').disabled = true; $('sheet').disabled = true;
  try {
    const response = await fetch('/api/calculate' + (sheet ? '?sheet=' + encodeURIComponent(sheet) : ''), {method:'POST', body:file});
    const data = await response.json(); if (!response.ok) throw Error(data.error);
    if (!sheet) {
      $('sheet').replaceChildren(...data.sheets.map(name => new Option(name, name)));
      $('sheetbox').hidden = data.sheets.length < 2;
    }
    if (data.selection_required) { $('status').textContent = 'Wybierz arkusz wyceny.'; $('sheet').selectedIndex = -1; return; }
    result = data; $('status').textContent = 'Gotowe. Kopiuj wartości przy procesach, dodatkach i tworzywach.'; render();
  } catch (e) { $('status').textContent = e.message || 'Nie udało się odczytać pliku.'; }
  finally { $('file').disabled = false; $('sheet').disabled = false; }
}
function render() {
  if (!result) return;
  $('results').hidden = false; $('project').textContent = result.project || 'Twoja wycena';
  $('meta').textContent = `${result.quantity} szt. · arkusz ${result.sheet} · ${result.details.length} pozycji`;
  $('rows').replaceChildren(...result.processes.map(p => {
    const row = document.createElement('tr');
    for (const text of [p.name, duration(p.hours), decimal(p.per_item)]) {
      const cell = document.createElement('td'); cell.textContent = text; row.append(cell);
    }
    const cell = document.createElement('td'), button = document.createElement('button');
    button.textContent = 'Kopiuj'; button.setAttribute('aria-label', `Kopiuj czas ${p.name}`);
    button.onclick = () => copy(decimal(p.per_item)); cell.append(button); row.append(cell); return row;
  }));
  $('setup').textContent = `Rozruch produkcji: ${duration(result.setup_hours)} — uwzględniony w MANUALNE. Łączne czasy wyświetlamy z dokładnością do sekundy; obliczenia zachowują pełną dokładność.`;
  $('rounding').textContent = result.processes.some(p => p.per_item > 0 && Number(p.per_item.toFixed(Number($('precision').value))) === 0) ? 'Uwaga: mały czas zaokrągla się do zera. Zwiększ liczbę miejsc po przecinku.' : '';
  $('details').replaceChildren(...result.details.map(d => {const p=document.createElement('p'); p.textContent=`Wiersz ${d.row} · ${d.label} → ${d.group} · ${duration(d.hours)}`; return p;}));
  const materialRow = (m, summary = false) => {
    const row = document.createElement('tr');
    for (const text of [m.description + (summary ? ` (${m.count} poz.)` : ''), decimal(m.total_m2), decimal(m.per_item)]) {
      const cell = document.createElement('td'); cell.textContent = text; row.append(cell);
    }
    const cell = document.createElement('td'), button = document.createElement('button');
    button.textContent = 'Kopiuj'; button.setAttribute('aria-label', `Kopiuj zużycie ${m.description}`);
    button.onclick = () => copy(decimal(m.per_item)); cell.append(button); row.append(cell);
    return row;
  };
  $('materials').replaceChildren(...result.materials.map(m => materialRow(m)));
  $('materialtotals').replaceChildren(...result.material_totals.map(m => materialRow(m, true)));
  $('materialsummary').hidden = !result.material_totals.length;
  $('copymaterialtotals').disabled = !result.material_totals.length;
  $('materialstatus').textContent = result.material_warnings.length
    ? result.material_warnings.join('\n')
    : result.materials.length ? 'Zużycie na sztukę = m² łącznie / nakład z K4.' : 'Brak opisanych bloków tworzyw.';
  $('materialrounding').textContent = result.materials.some(m => m.per_item > 0 && Number(m.per_item.toFixed(Number($('precision').value))) === 0)
    ? 'Małe zużycie zaokrągla się do zera. Zwiększ liczbę miejsc po przecinku.' : '';
  $('copymaterials').disabled = !result.materials.length;
  $('additions').replaceChildren(...result.additions.map(a => {
    const row = document.createElement('tr');
    const label = [a.specification, a.description].filter(Boolean).join(' · ');
    for (const text of [label, `${a.section} / ${a.unit}`, decimal(a.per_item)]) {
      const cell = document.createElement('td'); cell.textContent = text; row.append(cell);
    }
    const cell = document.createElement('td'), button = document.createElement('button');
    button.textContent = 'Kopiuj'; button.setAttribute('aria-label', `Kopiuj ilość ${label}`);
    button.onclick = () => copy(decimal(a.per_item)); cell.append(button); row.append(cell);
    return row;
  }));
  $('additionstatus').textContent = result.addition_warnings.length ? result.addition_warnings.join('\n')
    : result.additions.length ? 'Ilości na jedną sztukę wyrobu. Pianka: m² / szt. wyrobu.' : 'Brak opisanych dodatków.';
  $('additionrounding').textContent = result.additions.some(a => a.per_item > 0 && Number(a.per_item.toFixed(Number($('precision').value))) === 0)
    ? 'Mała ilość dodatku zaokrągla się do zera. Zwiększ liczbę miejsc po przecinku.' : '';
  $('copyadditions').disabled = !result.additions.length;
}
$('file').onchange = () => {
  file = $('file').files[0]; $('results').hidden = true; $('sheetbox').hidden = true; result = null;
  if (!file) return;
  $('filename').textContent = file.name;
  if (!/\.xlsx$/i.test(file.name) || file.size > 10 * 1024 * 1024) { $('status').textContent = 'Wybierz plik XLSX do 10 MB.'; return; }
  calculate();
};
$('sheet').onchange = () => calculate($('sheet').value);
$('precision').onchange = render;
$('copyall').onclick = () => result && copy('Proces\tGodziny / szt.\n' + result.processes.map(p => `${p.name}\t${decimal(p.per_item)}`).join('\n'));
$('copymaterials').onclick = () => result && copy('Tworzywo i opis\tm² / szt.\n' + result.materials.map(m => `${m.description.replace(/[\t\r\n]+/g, ' ')}\t${decimal(m.per_item)}`).join('\n'));
$('copymaterialtotals').onclick = () => result && copy('Tworzywo\tm² łącznie\tm² / szt.\n' + result.material_totals.map(m => `${m.description.replace(/[\t\r\n]+/g, ' ')}\t${decimal(m.total_m2)}\t${decimal(m.per_item)}`).join('\n'));
$('copyadditions').onclick = () => result && copy('Specyfikacja i opis\tSekcja\tJednostka\tIlość / szt. wyrobu\n' + result.additions.map(a => `${[a.specification, a.description].filter(Boolean).join(' · ').replace(/[\t\r\n]+/g, ' ')}\t${a.section}\t${a.unit}\t${decimal(a.per_item)}`).join('\n'));
