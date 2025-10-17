let worksheet, settings, root, _timer;

async function init() {
  try {
    await tableau.extensions.initializeAsync();
  } catch (e) {
    console.error("Tableau initialize failed (okay in local file preview):", e);
  }

  const isInTableau = !!(tableau && tableau.extensions && tableau.extensions.dashboardContent);

  if (!isInTableau) {
    // Not running inside Tableau (e.g., local browser preview)
    document.querySelector("#controls").insertAdjacentHTML(
      "beforeend",
      '<span style="margin-left:12px;opacity:.7">Preview mode — open from Tableau to bind data.</span>'
    );
    return;
  }

  const dashboard = tableau.extensions.dashboardContent.dashboard;

  settings = tableau.extensions.settings.getAll();
  const wsSelect = document.getElementById('wsSelect');
  dashboard.worksheets.forEach(w => {
    const opt = document.createElement('option');
    opt.value = w.name;
    opt.textContent = w.name;
    wsSelect.appendChild(opt);
  });

  if (settings.worksheet) wsSelect.value = settings.worksheet;
  document.getElementById('idField').value = settings.idField || '';
  document.getElementById('parentField').value = settings.parentField || '';
  document.getElementById('labelField').value = settings.labelField || '';

  document.getElementById('save').onclick = async () => {
    tableau.extensions.settings.set('worksheet', wsSelect.value);
    tableau.extensions.settings.set('idField', document.getElementById('idField').value);
    tableau.extensions.settings.set('parentField', document.getElementById('parentField').value);
    tableau.extensions.settings.set('labelField', document.getElementById('labelField').value);
    await tableau.extensions.settings.saveAsync();
    await render();
  };

  if (settings.worksheet && settings.idField && settings.parentField && settings.labelField) {
    await render();
  }
}

async function render() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  worksheet = dashboard.worksheets.find(w => w.name === tableau.extensions.settings.get('worksheet'));
  if (!worksheet) return;

  worksheet.removeEventListener(tableau.TableauEventType.FilterChanged, onSheetUpdated);
  worksheet.removeEventListener(tableau.TableauEventType.MarkSelectionChanged, onSheetUpdated);
  worksheet.addEventListener(tableau.TableauEventType.FilterChanged, onSheetUpdated);
  worksheet.addEventListener(tableau.TableauEventType.MarkSelectionChanged, onSheetUpdated);

  const data = await getData();
  const { idField, parentField, labelField } = tableau.extensions.settings.getAll();
  const { nodes, links } = transformToHierarchy(data, idField, parentField, labelField);
  draw(nodes, links, idField, labelField);
}

async function getData() {
  const rdr = await worksheet.getSummaryDataReaderAsync();
  const dt = await rdr.getAllPagesAsync();
  await rdr.releaseAsync();

  const cols = dt.columns.map(c => c.fieldName);
  return dt.data.map(row => Object.fromEntries(row.map((cell, i) => [cols[i], cell.formattedValue ?? cell.value])));
}

function transformToHierarchy(rows, idKey, parentKey, labelKey) {
  const clean = rows.filter(r => r[idKey]);
  const ids = new Set(clean.map(r => String(r[idKey])));
  const pruned = clean.filter(r => !r[parentKey] || ids.has(String(r[parentKey])));

  const stratify = d3.stratify()
    .id(d => String(d[idKey]))
    .parentId(d => (d[parentKey] ? String(d[parentKey]) : null));

  root = stratify(pruned);
  root.each(d => { d.label = d.data[labelKey] || d.id; });

  const treeLayout = d3.tree().nodeSize([26, 200]); // [nodeHeight, nodeWidth]
  treeLayout(root);

  const nodes = root.descendants();
  const links = root.links();
  return { nodes, links };
}

function draw(nodes, links, idKey, labelKey) {
  const container = document.getElementById('viz');
  const width = container.clientWidth || 1200;
  const height = Math.max(600, nodes.length * 28);

  d3.select('#viz').selectAll('*').remove();

  const svg = d3.select('#viz').append('svg')
    .attr('width', width)
    .attr('height', height);

  const zoom = d3.zoom().on('zoom', (event) => g.attr('transform', event.transform));
  svg.call(zoom);

  const g = svg.append('g').attr('transform', 'translate(40,40)');

  // Links (teal on dark)
  g.append('g')
    .selectAll('path')
    .data(links)
    .join('path')
    .attr('fill', 'none')
    .attr('stroke', '#2dd4bf') // teal-400
    .attr('stroke-opacity', 0.7)
    .attr('stroke-width', 1.2)
    .attr('d', d => `
      M${d.source.y},${d.source.x}
      C${(d.source.y + d.target.y)/2},${d.source.x}
       ${(d.source.y + d.target.y)/2},${d.target.x}
       ${d.target.y},${d.target.x}`);

  // Nodes
  const node = g.append('g')
    .selectAll('g.node')
    .data(nodes)
    .join('g')
    .attr('class', 'node')
    .attr('transform', d => `translate(${d.y},${d.x})`)
    .on('click', async (event, d) => {
      const idVal = d.data[idKey];
      await worksheet.selectMarksByValueAsync(
        [{ fieldName: idKey, value: idVal }],
        tableau.SelectionUpdateType.Replace
      );
    });

  node.append('circle')
    .attr('r', 5)
    .attr('stroke', '#a3a3a3')
    .attr('fill', '#111'); // match bg

  node.append('text')
    .attr('dy', '0.32em')
    .attr('x', 10)
    .text(d => d.label)
    .style('font-family', 'system-ui, sans-serif')
    .style('font-size', '12px')
    .style('fill', '#e5e5e5'); // light text
}

function onSheetUpdated() {
  clearTimeout(_timer);
  _timer = setTimeout(render, 200);
}

init();
