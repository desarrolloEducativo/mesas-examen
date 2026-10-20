'use strict';

function esMesaCompartidaRadiologia(mesa) {
  return mesa.carrera === 'Radiología' && mesa.anio === '1' && mesa.resolucion === '5308/24';
}

function obtenerResolucionesMesa(mesa) {
  return esMesaCompartidaRadiologia(mesa) ? ['5308/24', '1369/25'] : [mesa.resolucion];
}

function crearTarjetaMesa(mesa, indice) {
  function elemento(tipo, clase, texto) {
    const nodo = document.createElement(tipo);
    if (clase) nodo.className = clase;
    if (texto !== undefined) nodo.textContent = texto;
    return nodo;
  }

  function dato(lista, etiqueta, contenido, clase) {
    const fila = elemento('div', clase);
    const valor = elemento('dd');
    if (typeof contenido === 'string') valor.textContent = contenido;
    else valor.append(contenido);
    fila.append(elemento('dt', '', etiqueta), valor);
    lista.append(fila);
    return valor;
  }

  const tarjeta = elemento('article', 'mesa');
  const titulo = elemento('h2', '', mesa.materia);
  titulo.id = 'mesa-titulo-' + indice;
  tarjeta.setAttribute('aria-labelledby', titulo.id);
  tarjeta.append(titulo);

  const horario = elemento('dl', 'mesa-horario');
  const fecha = elemento('time');
  fecha.setAttribute('datetime', mesa.fecha);
  fecha.append(
    elemento('span', 'mesa-dia', mesa.dia),
    elemento('span', 'mesa-fecha', formatoFecha.format(new Date(mesa.fecha + 'T12:00:00Z')))
  );
  dato(horario, 'Fecha del examen', fecha);
  const hora = elemento('time', 'mesa-hora', mesa.hora + ' hs');
  hora.setAttribute('datetime', mesa.fecha + 'T' + mesa.hora);
  const valorHorario = dato(horario, 'Horario', hora);
  if (mesa.turno.trim()) valorHorario.append(elemento('span', 'mesa-turno', 'Turno ' + mesa.turno.trim()));
  tarjeta.append(horario);

  const plan = elemento('dl', 'mesa-plan');
  dato(plan, 'Carrera', mesa.carrera, 'mesa-carrera');
  if (mesa.anio) dato(plan, 'Año', /^\d+$/.test(mesa.anio) ? mesa.anio + '.º año' : mesa.anio);
  const resoluciones = obtenerResolucionesMesa(mesa);
  if (mesa.resolucion) dato(plan, resoluciones.length > 1 ? 'Resoluciones' : 'Resolución', resoluciones.join(' y '));
  tarjeta.append(plan);

  const lugar = elemento('dl', 'mesa-lugar');
  for (const [campo, etiqueta] of [['sede', 'Sede'], ['aula', 'Aula'], ['comision', 'Comisión']]) {
    if (mesa[campo] && mesa[campo].trim()) dato(lugar, etiqueta, mesa[campo].trim());
  }
  if (lugar.children.length) tarjeta.append(lugar);

  const tribunal = elemento('section', 'mesa-tribunal');
  tribunal.append(elemento('h3', '', 'Docentes del tribunal'));
  const personas = elemento('dl');
  const nombres = [mesa.docente, ...obtenerVocales(mesa)];
  nombres.forEach((nombre, posicion) => {
    const valor = dato(personas, posicion === 0 ? 'Titular' : 'Vocal ' + posicion, nombre || 'No informado', 'mesa-persona');
    if (!nombre) valor.className = 'mesa-sin-dato';
  });
  tribunal.append(personas);
  tarjeta.append(tribunal);
  return tarjeta;
}
