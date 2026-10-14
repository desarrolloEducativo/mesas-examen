const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const context = vm.createContext({});
vm.runInContext(fs.readFileSync('appscript/Code.gs', 'utf8'), context);
const base = JSON.parse(fs.readFileSync('planes_estudio.json', 'utf8'));
const headers = ['Carrera','Resolución','Año','Código materia','Materia','Regularizadas para cursar','Aprobadas para cursar','Regularizadas para rendir','Aprobadas para rendir','Revisión','Observaciones','ID resolución','ID materia'];
const tipos = ['reg_cursar','apr_cursar','reg_rendir','apr_rendir'];
function hojasFuente() {
  return base.carreras.flatMap(c => c.resoluciones.map(r => ({
    nombre: r.codigo,
    filas: [headers, ...r.materias.map(m => [c.nombre, `${r.numero_resolucion}/${r.anio_resolucion}`, String(m.anio), String(m.codigo), m.nombre,
      ...tipos.map(t => m.correlatividades.filter(co => co.tipo === t).map(co => {
        const req = r.materias.find(x => x.id === co.materia_requerida_id);
        return `${req.codigo} · ${req.nombre}`;
      }).join(', ')), 'Revisada', '', String(r.id), String(m.id)])]
  })));
}
const original = JSON.stringify(base);
const sinCambios = context.construirPlanes_(base, hojasFuente());
assert.equal(sinCambios.revisadas, 401);
assert.equal(sinCambios.cambios.length, 0);
assert.equal(JSON.stringify(sinCambios.datos), original);
let hojas = hojasFuente();
const h = hojas.find(h => h.filas.length > 3 && h.filas[1][5] === '');
const seleccion = `${h.filas[2][3]} · ${h.filas[2][4]}, ${h.filas[3][3]} · ${h.filas[3][4]}`;
h.filas[1][5] = seleccion;
let actualizado = context.construirPlanes_(base, hojas);
assert.equal(actualizado.cambios.length, 1);
assert.equal(actualizado.datos.cantidades.correlatividades, base.cantidades.correlatividades + 2);
assert.equal(JSON.stringify(base), original);
h.filas[1][9] = 'Pendiente';
assert.equal(context.construirPlanes_(base, hojas).cambios.length, 0);
h.filas[1][9] = 'Revisada';
h.filas[1][5] = `${h.filas[1][3]} · ${h.filas[1][4]}`;
assert.throws(() => context.construirPlanes_(base, hojas), /sí misma/);
h.filas[1][5] = '999999 · DESCONOCIDA';
assert.throws(() => context.construirPlanes_(base, hojas), /desconocida/);
h.filas[1][5] = `${h.filas[2][3]} · ${h.filas[2][4]}, ${h.filas[2][3]} · ${h.filas[2][4]}`;
assert.throws(() => context.construirPlanes_(base, hojas), /repetida/);
hojas = hojasFuente();
const conRequisitos = hojas.find(h => h.filas.slice(1).some(f => f.slice(5, 9).some(Boolean)));
const fila = conRequisitos.filas.find(f => f[9] === 'Revisada' && f.slice(5, 9).some(Boolean));
const cantidadAnterior = fila.slice(5, 9).filter(Boolean).length;
fila.fill('', 5, 9);
assert.equal(context.construirPlanes_(base, hojas).cambios.length, 1);
assert(context.construirPlanes_(base, hojas).datos.cantidades.correlatividades < base.cantidades.correlatividades);
hojas = hojasFuente(); hojas.find(h => h.filas.length > 1).filas.pop();
assert.throws(() => context.construirPlanes_(base, hojas), /Falta la materia/);
hojas = hojasFuente(); const duplicada = hojas.find(h => h.filas.length > 1); duplicada.filas.push([...duplicada.filas[1]]);
assert.throws(() => context.construirPlanes_(base, hojas), /duplicada/);
const plan = { materias: [{id:1,codigo:'01',nombre:'LECTURA, ESCRITURA Y ORALIDAD'}, {id:2,codigo:'02',nombre:'OTRA'}] };
assert.deepEqual(Array.from(context.resolverSelecciones_('01 · LECTURA, ESCRITURA Y ORALIDAD, 02 · OTRA', plan, 3)), [1,2]);
assert.throws(() => context.resolverSelecciones_('01 · LECTURA, ESCRITURA Y ORALIDAD, basura', plan, 3), /desconocida/);
console.log('OK: filas y relaciones del catálogo, selecciones múltiples, nombres con comas, altas, bajas, pendientes, preservación y datos inválidos.');
const hojaUnica = hojasFuente().find(h => h.filas.length > 3);
const idUnico = hojaUnica.filas[1][11];
hojaUnica.filas[1][5] = `${hojaUnica.filas[2][3]} · ${hojaUnica.filas[2][4]}`;
const parcial = context.construirPlanes_(base, [hojaUnica], idUnico);
assert.equal(parcial.cambios.length, 1);
for (const carrera of base.carreras) for (const p of carrera.resoluciones) {
  if (String(p.id) !== idUnico) {
    const guardado = parcial.datos.carreras.flatMap(c => c.resoluciones).find(r => r.id === p.id);
    assert.equal(JSON.stringify(guardado), JSON.stringify(p));
  }
}
const incompleta = JSON.parse(JSON.stringify(hojaUnica)); incompleta.filas.pop();
assert.throws(() => context.construirPlanes_(base, [incompleta], idUnico), /Falta la materia/);
assert.throws(() => context.construirPlanes_(base, [hojaUnica], '999999'), /desconocida/);

let boton = 'NO', escrituras = 0, bloqueos = 0, liberaciones = 0;
const llamadas = [];
const ui = { ButtonSet: {YES_NO:'YES_NO'}, Button:{YES:'YES'}, alert(...args) {
  if (args.length === 3) { assert.equal(bloqueos, liberaciones); llamadas.push('confirmar'); return boton; }
  llamadas.push('mensaje');
}};
context.SpreadsheetApp = {getUi:()=>ui};
context.LockService = {getScriptLock:()=>({tryLock:()=>{bloqueos++; llamadas.push('lock');return true;},releaseLock:()=>{liberaciones++;llamadas.push('release');}})};
context.Utilities = {base64Encode:s=>Buffer.from(s).toString('base64'),Charset:{UTF_8:'utf8'}};
context.prepararPlanes_ = activa => ({...parcial,config:{owner:'owner',repo:'repo',branch:'main',path:'planes_estudio.json'},sha:'original',alcance:activa?'Activa':'Todas'});
context.githubPlanes_ = (config,method,payload) => {
  assert.equal(method,'put');assert.equal(payload.sha,'original');escrituras++;llamadas.push('put');return {commit:{html_url:'https://example.com/commit'}};
};
for (const respuesta of ['NO','CLOSE']) {boton=respuesta;context.publicarPlanesGithub();assert.equal(escrituras,0);assert.equal(bloqueos,0);}
boton='YES';llamadas.length=0;context.publicarPlanesGithub();assert.equal(escrituras,1);assert.deepEqual(llamadas,['confirmar','lock','put','release','mensaje']);
context.prepararPlanes_ = ()=>({...parcial,cambios:[]});context.publicarPlanesGithub();assert.equal(escrituras,1);
console.log('OK: actualización de un solo plan, preservación de otros planes y confirmación Sí/No/cierre antes del commit.');
const comentarioHoja = hojasFuente().find(h => h.filas.length > 1);
const idComentario = comentarioHoja.filas[1][11];
comentarioHoja.tieneComentarios = true;
comentarioHoja.idComentarios = idComentario;
comentarioHoja.comentarios = 'Primer comentario\nSegunda línea';
comentarioHoja.filas.slice(1).forEach(f => f[9] = 'Pendiente');
const conComentario = context.construirPlanes_(base,[comentarioHoja],idComentario);
assert.equal(conComentario.cambios.length,0);
assert.equal(conComentario.cambiosComentarios.length,1);
assert.equal(conComentario.datos.carreras.flatMap(c=>c.resoluciones).find(r=>String(r.id)===idComentario).comentarios,comentarioHoja.comentarios);
comentarioHoja.comentarios='';
const sinComentario=context.construirPlanes_(conComentario.datos,[comentarioHoja],idComentario);
assert.equal(sinComentario.cambiosComentarios.length,1);
delete comentarioHoja.tieneComentarios;
assert.equal(context.construirPlanes_(conComentario.datos,[comentarioHoja],idComentario).cambiosComentarios.length,0);
const vacio=base.carreras.flatMap(c=>c.resoluciones).find(r=>!r.materias.length);
const hojaVacia={nombre:vacio.codigo,filas:[headers],tieneComentarios:true,idComentarios:String(vacio.id),comentarios:'Plan en preparación'};
assert.equal(context.construirPlanes_(base,[hojaVacia],String(vacio.id)).cambiosComentarios.length,1);
console.log('OK: comentarios sin cambios en materias, eliminación, compatibilidad anterior y planes vacíos.');
const idPdf = String(base.carreras[0].resoluciones[0].id);
const enlacePdf = 'https://drive.google.com/file/d/ejemplo/view?usp=sharing';
const indicePdf = [['ID resolución', 'Link PDF'], [idPdf, enlacePdf]];
const conPdf = context.construirPlanes_(base, hojasFuente(), undefined, indicePdf);
assert.equal(conPdf.cambiosPdf.length, 1);
assert.equal(conPdf.datos.carreras[0].resoluciones[0].pdf, enlacePdf);
assert.equal(JSON.stringify(base), original);
assert.equal(context.construirPlanes_(conPdf.datos, hojasFuente(), undefined, indicePdf).cambiosPdf.length, 0);
assert.equal(context.construirPlanes_(conPdf.datos, hojasFuente()).datos.carreras[0].resoluciones[0].pdf, enlacePdf);
assert.equal(context.construirPlanes_(conPdf.datos, hojasFuente(), undefined, [['ID resolución','Link PDF'],[idPdf,'']]).datos.carreras[0].resoluciones[0].pdf, '');
assert.throws(() => context.construirPlanes_(base, hojasFuente(), undefined, [['Link PDF'],[enlacePdf]]), /ID resolución/);
assert.throws(() => context.construirPlanes_(base, hojasFuente(), undefined, [['ID resolución','Link PDF'],[idPdf,'javascript:alert(1)']]), /HTTPS/);
assert.throws(() => context.construirPlanes_(base, hojasFuente(), undefined, [...indicePdf,indicePdf[1]]), /duplicada/);
const otroId = String(base.carreras[1].resoluciones[0].id);
assert.equal(context.construirPlanes_(base, hojasFuente().filter(h => h.filas[1] && h.filas[1][11] === otroId), otroId, indicePdf).cambiosPdf.length, 0);
console.log('OK: PDF de Drive, eliminación, compatibilidad anterior, alcance y enlaces inválidos.');
