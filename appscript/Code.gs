/** Script vinculado a la planilla. Configurar GITHUB_TOKEN en Propiedades del script. */
const CONFIG_PLANES = {
  owner: 'desarrolloEducativo',
  repo: 'mesas-examen',
  branch: 'main',
  path: 'planes_estudio.json'
};
const TIPOS_PLANES = {
  'Regularizadas para cursar': 'reg_cursar',
  'Aprobadas para cursar': 'apr_cursar',
  'Regularizadas para rendir': 'reg_rendir',
  'Aprobadas para rendir': 'apr_rendir'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Planes de estudio')
    .addItem('Validar y ver resumen', 'validarPlanes')
    .addItem('Descargar JSON', 'descargarPlanesJson')
    .addItem('Agregar campo de comentarios a las hojas', 'agregarCamposComentarios')
    .addSeparator()
    .addItem('Actualizar todas las resoluciones', 'actualizarTodasLasResoluciones')
    .addItem('Actualizar resolución de la pestaña activa', 'actualizarResolucionActiva')
    .addToUi();
}

function configurarPlanes_() {
  const p = PropertiesService.getScriptProperties();
  const config = Object.assign({}, CONFIG_PLANES);
  ['owner', 'repo', 'branch', 'path'].forEach(k => {
    config[k] = p.getProperty('GITHUB_' + k.toUpperCase()) || config[k];
  });
  config.token = p.getProperty('GITHUB_TOKEN');
  if (!config.token) throw new Error('Falta GITHUB_TOKEN en las Propiedades del script.');
  return config;
}

function githubPlanes_(config, method, payload) {
  const url = 'https://api.github.com/repos/' + encodeURIComponent(config.owner) + '/' +
    encodeURIComponent(config.repo) + '/contents/' + config.path.split('/').map(encodeURIComponent).join('/') +
    (method === 'get' ? '?ref=' + encodeURIComponent(config.branch) : '');
  const opciones = {
    method, muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + config.token, Accept: 'application/vnd.github+json' }
  };
  if (payload) {
    opciones.contentType = 'application/json';
    opciones.payload = JSON.stringify(payload);
  }
  const respuesta = UrlFetchApp.fetch(url, opciones);
  const codigo = respuesta.getResponseCode();
  if (codigo < 200 || codigo >= 300) {
    if (codigo === 404) throw new Error('No se encontró el JSON en GitHub. Publicá primero planes_estudio.json y revisá repositorio, rama y ruta.');
    if (codigo === 409 || codigo === 422) throw new Error('GitHub no aceptó la actualización. Puede haber cambios simultáneos o una regla de la rama. Volvé a validar antes de publicar.');
    throw new Error('GitHub respondió HTTP ' + codigo + '. Revisá el token, sus permisos y la configuración.');
  }
  return JSON.parse(respuesta.getContentText());
}

function prepararPlanes_(hojaActiva) {
  const config = configurarPlanes_();
  const archivo = githubPlanes_(config, 'get');
  if (!archivo.content || archivo.encoding !== 'base64') throw new Error('GitHub no devolvió el contenido del JSON en el formato esperado.');
  const base = JSON.parse(Utilities.newBlob(Utilities.base64Decode(archivo.content.replace(/\s/g, ''))).getDataAsString('UTF-8'));
  const hojas = (hojaActiva ? [hojaActiva] : SpreadsheetApp.getActiveSpreadsheet().getSheets())
    .filter(h => !['planes', 'instrucciones'].includes(h.getName().toLowerCase()))
    .map(h => {
      const filas = h.getDataRange().getDisplayValues();
      const tieneComentarios = String((filas[0] || [])[16] || '').trim() === 'Comentarios de la resolución';
      return { nombre: h.getName(), filas, tieneComentarios,
        comentarios: tieneComentarios ? String((filas[1] || [])[16] || '') : undefined,
        idComentarios: tieneComentarios ? String((filas[4] || [])[16] || '').trim() : undefined };
    });
  let idResolucion;
  if (hojaActiva) {
    if (!hojas.length) throw new Error('La pestaña activa no es una hoja de resolución.');
    const cabecera = hojas[0].filas[0] || [];
    const columna = cabecera.indexOf('ID resolución');
    if (columna < 0) throw new Error('La pestaña activa no tiene ID resolución.');
    const ids = [...new Set(hojas[0].filas.slice(1).map(f => String(f[columna] || '').trim()).filter(Boolean))];
    if (hojas[0].idComentarios && !ids.includes(hojas[0].idComentarios)) ids.push(hojas[0].idComentarios);
    if (!ids.length) throw new Error('Esta resolución no tiene materias cargadas para actualizar.');
    if (ids.length !== 1) throw new Error('La pestaña activa mezcla resoluciones.');
    idResolucion = ids[0];
  }
  const indice = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(h => h.getName().toLowerCase() === 'planes');
  const resultado = construirPlanes_(base, hojas, idResolucion, indice ? indice.getDataRange().getDisplayValues() : undefined);
  return Object.assign(resultado, { config, sha: archivo.sha, alcance: hojaActiva ? 'Resolución de la pestaña «' + hojaActiva.getName() + '»' : 'Todas las resoluciones' });
}

/** Función pura: valida toda la planilla y modifica únicamente correlativas revisadas. */
function construirPlanes_(base, hojas, idResolucion, filasPlanes) {
  const datos = JSON.parse(JSON.stringify(base));
  if (!Array.isArray(datos.carreras)) throw new Error('El JSON base no contiene carreras.');
  const planes = new Map();
  let siguienteId = 1;
  datos.carreras.forEach(c => c.resoluciones.forEach(r => {
    planes.set(String(r.id), r);
    r.materias.forEach(m => (m.correlatividades || []).forEach(co => {
      siguienteId = Math.max(siguienteId, Number(co.id || 0) + 1);
    }));
  }));
  const vistos = new Set();
  if (idResolucion !== undefined && !planes.has(String(idResolucion))) throw new Error('Resolución desconocida: ' + idResolucion);
  const errores = [];
  const cambios = [];
  const cambiosComentarios = [];
  const cambiosPdf = [];
  if (filasPlanes) {
    const cabecera = (filasPlanes[0] || []).map(v => String(v).trim());
    const colPdf = cabecera.indexOf('Link PDF');
    if (colPdf >= 0) {
      const colId = cabecera.indexOf('ID resolución');
      if (colId < 0) throw new Error('Planes: falta la columna ID resolución para identificar cada PDF.');
      if (cabecera.lastIndexOf('Link PDF') !== colPdf || cabecera.lastIndexOf('ID resolución') !== colId) throw new Error('Planes: hay encabezados duplicados.');
      const vistosPdf = new Set();
      filasPlanes.slice(1).forEach((fila, i) => {
        const id = String(fila[colId] || '').trim();
        const pdf = String(fila[colPdf] || '').trim();
        if (!id && !pdf) return;
        if (!planes.has(id)) { errores.push('Planes, fila ' + (i + 2) + ': ID resolución desconocido.'); return; }
        if (vistosPdf.has(id)) { errores.push('Planes, fila ' + (i + 2) + ': resolución duplicada.'); return; }
        vistosPdf.add(id);
        if (idResolucion !== undefined && id !== String(idResolucion)) return;
        if (pdf && !/^https:\/\/[^\s/?#]+(?:[/?#][^\s]*)?$/i.test(pdf)) { errores.push('Planes, fila ' + (i + 2) + ': Link PDF debe ser un enlace HTTPS completo.'); return; }
        const plan = planes.get(id);
        if (String(plan.pdf || '') !== pdf) { plan.pdf = pdf; cambiosPdf.push({ resolucion: id, nombre: plan.codigo }); }
      });
    }
  }
  const comentariosVistos = new Set();
  let revisadas = 0, pendientes = 0;
  const requerido = ['ID resolución', 'ID materia', 'Código materia', 'Revisión', ...Object.keys(TIPOS_PLANES)];
  hojas.forEach(hoja => {
    const cabecera = (hoja.filas[0] || []).map(x => String(x).trim());
    if (!requerido.every(x => cabecera.includes(x))) {
      errores.push(hoja.nombre + ': faltan encabezados requeridos.'); return;
    }
    if (requerido.some(x => cabecera.indexOf(x) !== cabecera.lastIndexOf(x))) {
      errores.push(hoja.nombre + ': hay encabezados duplicados.'); return;
    }
    const columna = nombre => cabecera.indexOf(nombre);
    if (hoja.tieneComentarios) {
      const id = String(hoja.idComentarios || '').trim();
      const planComentarios = planes.get(id);
      if (!planComentarios || (idResolucion !== undefined && id !== String(idResolucion))) {
        errores.push(hoja.nombre + ': ID de comentarios inválido.');
      } else if (comentariosVistos.has(id)) {
        errores.push(hoja.nombre + ': comentarios duplicados para la resolución.');
      } else {
        comentariosVistos.add(id);
        const valor = String(hoja.comentarios || '');
        if (String(planComentarios.comentarios || '') !== valor) {
          planComentarios.comentarios = valor;
          cambiosComentarios.push({ resolucion: id, nombre: planComentarios.codigo });
        }
      }
    }
    let idHoja;
    hoja.filas.slice(1).forEach((fila, i) => {
      const contexto = hoja.nombre + ', fila ' + (i + 2);
      const idPlan = String(fila[columna('ID resolución')] || '').trim();
      const idMateria = String(fila[columna('ID materia')] || '').trim();
      // Las columnas auxiliares y el aviso de planes vacíos no son registros.
      if (!idPlan && !idMateria) {
        if (fila.slice(1, 13).some(v => String(v).trim())) errores.push(contexto + ': fila con datos pero sin identificadores.');
        return;
      }
      const plan = planes.get(idPlan);
      if (idResolucion !== undefined && idPlan !== String(idResolucion)) {
        errores.push(contexto + ': la fila no pertenece a la resolución seleccionada.'); return;
      }
      const materia = plan && plan.materias.find(m => String(m.id) === idMateria);
      if (!materia) { errores.push(contexto + ': materia o resolución desconocida.'); return; }
      if (idHoja && idHoja !== idPlan) { errores.push(contexto + ': la hoja mezcla resoluciones.'); return; }
      idHoja = idPlan;
      if (hoja.tieneComentarios && idPlan !== String(hoja.idComentarios)) {
        errores.push(contexto + ': el ID del campo comentarios no coincide con la resolución de las materias.'); return;
      }
      const clave = idPlan + ':' + idMateria;
      if (vistos.has(clave)) { errores.push(contexto + ': materia duplicada.'); return; }
      vistos.add(clave);
      if (String(materia.codigo) !== String(fila[columna('Código materia')]).trim()) {
        errores.push(contexto + ': el código no coincide con el ID de materia.'); return;
      }
      const revision = String(fila[columna('Revisión')] || '').trim().toLowerCase();
      if (revision === 'pendiente' || !revision) { pendientes++; return; }
      if (revision !== 'revisada') { errores.push(contexto + ': Revisión debe ser Pendiente o Revisada.'); return; }
      revisadas++;
      const correlativas = [];
      Object.keys(TIPOS_PLANES).forEach(encabezado => {
        try {
          const ids = resolverSelecciones_(fila[columna(encabezado)], plan, materia.id);
          ids.forEach(id => {
            const tipo = TIPOS_PLANES[encabezado];
            const anterior = (materia.correlatividades || []).find(co => co.tipo === tipo && String(co.materia_requerida_id) === String(id));
            correlativas.push(anterior || { id: siguienteId++, tipo, materia_id: materia.id, materia_requerida_id: id });
          });
        } catch (e) { errores.push(contexto + ', ' + encabezado + ': ' + e.message); }
      });
      const firma = lista => lista.map(co => co.tipo + ':' + co.materia_requerida_id).sort().join('|');
      if (firma(materia.correlatividades || []) !== firma(correlativas)) {
        materia.correlatividades = correlativas;
        cambios.push({ resolucion: idPlan, materia: materia.nombre });
      }
    });
  });
  // Evita interpretar una planilla incompleta como un catálogo completo.
  planes.forEach(p => {
    if (idResolucion !== undefined && String(p.id) !== String(idResolucion)) return;
    p.materias.forEach(m => {
    if (!vistos.has(String(p.id) + ':' + m.id)) errores.push('Falta la materia ' + m.codigo + ' del plan ' + p.codigo + '.');
    });
  });
  if (errores.length) throw new Error(errores.slice(0, 25).join('\n') + (errores.length > 25 ? '\n… y ' + (errores.length - 25) + ' errores más.' : ''));
  if (cambios.length || cambiosComentarios.length || cambiosPdf.length) {
    datos.exportado_en = new Date().toISOString();
    datos.cantidades = Object.assign({}, datos.cantidades, {
      carreras: datos.carreras.length,
      resoluciones: planes.size,
      materias: [...planes.values()].reduce((n,p) => n + p.materias.length, 0),
      correlatividades: [...planes.values()].reduce((n,p) => n + p.materias.reduce((s,m) => s + (m.correlatividades || []).length, 0), 0)
    });
  }
  return { datos, cambios, cambiosComentarios, cambiosPdf, revisadas, pendientes, json: JSON.stringify(datos, null, 2) + '\n' };
}

function resolverSelecciones_(valor, plan, idMateria) {
  const texto = String(valor || '').trim();
  if (!texto) return [];
  // Reconoce el texto completo de cada chip, incluso nombres que contienen comas.
  const opciones = plan.materias.map(m => ({ m, etiqueta: m.codigo + ' · ' + m.nombre }))
    .sort((a,b) => b.etiqueta.length - a.etiqueta.length);
  const ids = [];
  let resto = texto;
  while (resto) {
    const opcion = opciones.find(o => resto === o.etiqueta || resto.startsWith(o.etiqueta + ',') || resto.startsWith(o.etiqueta + '\n'));
    if (!opcion) throw new Error('Selección desconocida o de otra resolución: ' + resto.slice(0, 90));
    if (String(opcion.m.id) === String(idMateria)) throw new Error('Una materia no puede ser correlativa de sí misma.');
    if (ids.includes(opcion.m.id)) throw new Error('Materia repetida: ' + opcion.m.codigo);
    ids.push(opcion.m.id);
    resto = resto.slice(opcion.etiqueta.length).replace(/^[,\n]\s*/, '').trim();
  }
  return ids;
}

function resumenPlanes_(r) {
  return r.revisadas + ' materias revisadas, ' + r.pendientes + ' pendientes.\n' +
    r.cambios.length + ' materias con cambios.\n' + (r.cambiosComentarios || []).length + ' resoluciones con cambios en comentarios.\n' + (r.cambiosPdf || []).length + ' resoluciones con cambios en PDF.\n' + r.datos.cantidades.correlatividades + ' correlatividades en total.';
}
function validarPlanes() {
  try { SpreadsheetApp.getUi().alert(resumenPlanes_(prepararPlanes_())); }
  catch (e) { SpreadsheetApp.getUi().alert('No se pudo validar:\n' + e.message); }
}
function descargarPlanesJson() {
  try {
    const r = prepararPlanes_();
    const b64 = Utilities.base64Encode(r.json, Utilities.Charset.UTF_8);
    const html = '<p>JSON validado. No se publicó en GitHub.</p><a download="planes_estudio.json" href="data:application/json;base64,' + b64 + '">Descargar planes_estudio.json</a>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(140), 'Descargar JSON');
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}
function actualizarTodasLasResoluciones() {
  publicarPlanesGithub();
}
function actualizarResolucionActiva() {
  publicarPlanesGithub(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
}
function publicarPlanesGithub(hojaActiva) {
  const ui = SpreadsheetApp.getUi();
  let lock;
  let mensaje;
  try {
    const r = prepararPlanes_(hojaActiva);
    if (!r.cambios.length && !(r.cambiosComentarios || []).length && !(r.cambiosPdf || []).length) { ui.alert('No hay cambios para publicar.\n' + resumenPlanes_(r)); return; }
    const confirmacion = ui.alert('Advertencia: actualizar resoluciones en GitHub',
      r.alcance + '\n\n' + resumenPlanes_(r) + '\n\n' +
      'Se actualizarán las correlativas de las filas Revisada. Sus celdas vacías eliminan los requisitos correspondientes. Las pendientes se conservan.\n\n' +
      'También se guardarán los comentarios de cada resolución seleccionada. Dejar ese campo vacío elimina el comentario anterior.\n\n' +
      'También se guardarán los enlaces de Link PDF en Planes. Una celda vacía elimina el enlace anterior.\n\n' +
      'Se creará un commit en ' + r.config.owner + '/' + r.config.repo + ', rama ' + r.config.branch + ', archivo ' + r.config.path + '.\n\n¿Confirmás la actualización?', ui.ButtonSet.YES_NO);
    if (confirmacion !== ui.Button.YES) return;
    // Los diálogos suspenden Apps Script y no conservan los locks.
    lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) { lock = undefined; throw new Error('Hay otra publicación en curso. Intentá nuevamente.'); }
    // Publicar exactamente el resultado confirmado. El SHA protege contra cambios posteriores.
    const respuesta = githubPlanes_(r.config, 'put', {
      message: 'Actualizar correlatividades desde Google Sheets: ' + r.alcance,
      content: Utilities.base64Encode(r.json, Utilities.Charset.UTF_8),
      sha: r.sha, branch: r.config.branch
    });
    mensaje = 'Publicado en GitHub.\n' + r.alcance + '\n' + resumenPlanes_(r) + '\n' + respuesta.commit.html_url;
  } catch (e) { mensaje = 'No se publicó el JSON:\n' + e.message; }
  finally { if (lock) lock.releaseLock(); }
  if (mensaje) ui.alert(mensaje);
}

/** Migra la planilla existente sin reemplazar materias ni correlativas ya editadas. */
function agregarCamposComentarios() {
  try {
    const config = configurarPlanes_();
    const archivo = githubPlanes_(config, 'get');
    const base = JSON.parse(Utilities.newBlob(Utilities.base64Decode(archivo.content.replace(/\s/g, ''))).getDataAsString('UTF-8'));
    const planes = base.carreras.flatMap(c => c.resoluciones.map(r => ({ plan: r, hoja: c.codigo + ' ' + r.numero_resolucion + '-' + r.anio_resolucion })));
    let agregados = 0;
    SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(h => !['planes', 'instrucciones'].includes(h.getName().toLowerCase())).forEach(h => {
      if (h.getRange('Q1').getDisplayValue() === 'Comentarios de la resolución') return;
      if (h.getRange('Q1:Q5').getDisplayValues().flat().some(Boolean)) throw new Error(h.getName() + ': Q1:Q5 contiene datos. No se reemplazaron.');
      const filas = h.getDataRange().getDisplayValues();
      const col = (filas[0] || []).indexOf('ID resolución');
      const ids = [...new Set(filas.slice(1).map(f => String(f[col] || '').trim()).filter(Boolean))];
      const entrada = ids.length === 1 ? planes.find(p => String(p.plan.id) === ids[0]) : planes.find(p => p.hoja === h.getName());
      if (!entrada || ids.length > 1) throw new Error('No se pudo identificar la resolución de ' + h.getName());
      h.getRange('Q1:Q5').setValues([['Comentarios de la resolución'], [entrada.plan.comentarios || ''], [''], ['ID resolución (comentarios)'], [String(entrada.plan.id)]]);
      h.setColumnWidth(17, 430);
      h.getRange('Q1:Q5').setWrap(true);
      h.getRange('Q2').setBackground('#FFF4CC');
      h.setRowHeight(2, Math.max(h.getRowHeight(2), 80));
      agregados++;
    });
    SpreadsheetApp.getUi().alert('Campos de comentarios agregados: ' + agregados + '. Editá Q2 en cada hoja. Las materias y correlativas se conservaron.');
  } catch (e) { SpreadsheetApp.getUi().alert('No se pudo completar la preparación:\n' + e.message); }
}
