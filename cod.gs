// Código JavaScript para Google Apps Script (cliente-side, pero extraído del HTML para formato .gs)
// Nota: Este código es para el lado del cliente en Apps Script. Para el servidor, necesitas definir las funciones llamadas en un archivo .gs separado.

document.addEventListener('DOMContentLoaded', () => {
  // Inicialización de Materialize
  M.AutoInit();
  M.updateTextFields();

  const SCRIPT_RUN = google.script.run.withFailureHandler(err => {
    M.toast({ html: 'Error del servidor: ' + err.message, classes: 'red darken-3' });
    console.error(err);
    document.querySelectorAll('button').forEach(btn => btn.disabled = false);
  });

  // --- Utilidades ---
  function ocultarTodas() {
    document.querySelectorAll('.page-content').forEach(p => p.classList.remove('active-page'));
  }

  // --- Autocomplete ---
  window.initAutocomplete = function (inputId, hiddenId, serverFunc, fetchDetails = false) {
    const elem = document.getElementById(inputId);
    if (!elem) return;
    const instance = M.Autocomplete.getInstance(elem);
    if (instance) instance.destroy();

    google.script.run
      .withFailureHandler(err => {
        console.error("Error en " + serverFunc + ": ", err);
        M.toast({ html: 'Error cargando datos de ' + serverFunc, classes: 'red' });
      })
      .withSuccessHandler(data => {
        if (!Array.isArray(data)) return;
        const opciones = {};
        const detalles = {};
        data.forEach(o => {
          const key = o.direccion + ' (' + o.id + ')';
          opciones[key] = null;
          if (fetchDetails) detalles[o.id] = o;
        });
        M.Autocomplete.init(elem, {
          data: opciones,
          limit: 10,
          minLength: 1,
          onAutocomplete: val => {
            const match = val.match(/\(([^)]+)\)$/);
            const selectedId = match ? match[1] : '';
            const hiddenElem = document.getElementById(hiddenId);
            if (hiddenElem) hiddenElem.value = selectedId;
            if (fetchDetails && selectedId && detalles[selectedId]) {
              const tablero = detalles[selectedId];
              document.getElementById('switchCurrentDireccion').textContent = tablero.direccion;
              document.getElementById('inputSwitch').value = tablero.switch || '';
              const tipoSel = document.getElementById('tipoConectividadSwitch');
              if (tipoSel) {
                tipoSel.value = (tablero.tipoConectividad || '').toUpperCase();
                M.FormSelect.init(tipoSel);
              }
              const provSel = document.getElementById('proveedorConectividadSwitch');
              const inputDetalle = document.getElementById('detalleProveedorSwitch');
              const divDetalle = document.querySelector('.detalle-proveedor-switch');
              if (provSel) {
                const provPlanilla = (tablero.proveedorConectividad || '').toUpperCase();
                const existeOpcion = Array.from(provSel.options).some(opt => opt.value === provPlanilla);
                if (existeOpcion) {
                  provSel.value = provPlanilla;
                  if (divDetalle) divDetalle.style.display = 'none';
                } else if (provPlanilla !== "") {
                  provSel.value = "OTRA";
                  if (inputDetalle) inputDetalle.value = provPlanilla;
                  if (divDetalle) divDetalle.style.display = 'block';
                }
                M.FormSelect.init(provSel);
              }
              M.updateTextFields();
              M.toast({ html: 'Datos cargados', classes: 'blue' });
            }
          }
        });
      })[serverFunc]();
  };

  // --- Cambio de proveedor ---
  function handleProveedorChange(selectId, detailSelector) {
    const selectElement = document.getElementById(selectId);
    if (!selectElement) return;
    selectElement.addEventListener('change', function () {
      const detailDiv = document.querySelector(detailSelector);
      const detailInput = detailDiv.querySelector('input');
      if (this.value === 'OTRA') {
        detailDiv.style.display = 'block';
        if (detailInput) detailInput.setAttribute('required', 'true');
      } else {
        detailDiv.style.display = 'none';
        if (detailInput) {
          detailInput.value = '';
          detailInput.removeAttribute('required');
        }
      }
      M.updateTextFields();
    });
  }

  // --- Geolocalización ---
  function obtenerUbicacion() {
    const latInput = document.getElementById('instalacionLatitud');
    const lngInput = document.getElementById('instalacionLongitud');
    const statusSpan = document.getElementById('ubicacionStatus');
    if (!latInput || !lngInput || !statusSpan) return;
    latInput.value = ''; lngInput.value = ''; M.updateTextFields();
    statusSpan.textContent = 'Buscando ubicación...';
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (position) => {
          const lat = position.coords.latitude.toFixed(6);
          const lng = position.coords.longitude.toFixed(6);
          latInput.value = lat; lngInput.value = lng;
          M.updateTextFields();
          statusSpan.textContent = 'Ubicación obtenida.';
          M.toast({ html: 'Ubicación registrada: ' + lat + ', ' + lng, classes: 'blue' });
        },
        (error) => {
          let msg = 'Error de ubicación: ';
          switch (error.code) {
            case error.PERMISSION_DENIED: msg += "Usuario denegó la solicitud."; break;
            case error.POSITION_UNAVAILABLE: msg += "Información de ubicación no disponible."; break;
            case error.TIMEOUT:
              msg += "Tiempo de espera agotado. Ingrese coordenadas manualmente.";
              statusSpan.textContent = 'Ingrese coordenadas manualmente en los campos.';
              break;
            default: msg += "Error desconocido."; break;
          }
          console.error("Error Geolocation:", msg, error);
          statusSpan.textContent = 'Error al obtener ubicación.';
          M.toast({ html: msg, classes: 'red darken-3' });
        }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
      );
    } else {
      statusSpan.textContent = 'Geolocalización no soportada.';
      M.toast({ html: 'Geolocalización no disponible.', classes: 'red darken-3' });
    }
  }

  window.showPage = function (pageId) {
    ocultarTodas();
    const page = document.getElementById(pageId + '-page');
    if (page) page.classList.add('active-page');

    if (pageId !== 'panel' && window._lastPage !== pageId) {
      if (SCRIPT_RUN.registrarAccesoUsuario) {
        SCRIPT_RUN.registrarAccesoUsuario('Página: ' + pageId.toUpperCase());
      }
      window._lastPage = pageId;
    }

    if (pageId === 'nuevo') {
      const formNuevo = document.getElementById('form-nuevo-tablero');
      if (formNuevo && !formNuevo._listenerAdded) {
        formNuevo.addEventListener('submit', e => {
          e.preventDefault();
          const btn = e.submitter; btn.disabled = true;
          const data = {
            direccion: document.getElementById('direccion').value,
            marcado: document.getElementById('marcado_select').value,
            observaciones: document.getElementById('observaciones').value
          };
          SCRIPT_RUN.withSuccessHandler(res => {
            M.toast({ html: res.success ? 'Tablero ' + res.id + ' creado.' : 'Error: ' + res.message, classes: res.success ? 'green darken-1' : 'red darken-3' });
            formNuevo.reset(); M.updateTextFields(); btn.disabled = false;
          }).crearNuevoTablero(data);
        });
        formNuevo._listenerAdded = true;
      }
    }

    // INSTALACIÓN
    if (pageId === 'instalacion') {
      initAutocomplete('autocompleteInstalacion', 'tableroInstalacionId', 'listarTablerosPedido');
      document.getElementById('instalacionLatitud').value = '';
      document.getElementById('instalacionLongitud').value = '';
      const statusSpan = document.getElementById('ubicacionStatus'); if (statusSpan) statusSpan.textContent = '';
      M.updateTextFields();
      const btnUbicacion = document.getElementById('btnObtenerUbicacion');
      if (btnUbicacion && !btnUbicacion._listenerAdded) {
        btnUbicacion.addEventListener('click', obtenerUbicacion);
        btnUbicacion._listenerAdded = true;
      }
      const formInstalacion = document.getElementById('form-instalacion');
      if (formInstalacion && !formInstalacion._listenerAdded) {
        formInstalacion.addEventListener('submit', e => {
          e.preventDefault();
          const btn = e.submitter; if (btn) btn.disabled = true;

          const id = document.getElementById('tableroInstalacionId')?.value || '';
          const latitud = document.getElementById('instalacionLatitud')?.value || '';
          const longitud = document.getElementById('instalacionLongitud')?.value || '';

          if (!id) {
            M.toast({ html: 'Debe seleccionar un Tablero.' });
            if (btn) btn.disabled = false;
            return;
          }

          if (!latitud || !longitud) {
            M.toast({
              html: 'Advertencia: Coordenadas vacías. Presione "Obtener Coordenadas" o ingrese manualmente.',
              classes: 'orange darken-3'
            });
          }

          const data = {
            id, latitud, longitud,
            observaciones: document.getElementById('obsInstalacion')?.value || ''
          };

          SCRIPT_RUN.withSuccessHandler(res => {
            M.toast({
              html: res.success ? 'Tablero ' + id + ' instalado con éxito.' : 'Error: ' + res.message,
              classes: res.success ? 'green darken-1' : 'red darken-3'
            });
            formInstalacion.reset();
            M.updateTextFields();
            if (btn) btn.disabled = false;
            showPage('panel');
          }).registrarInstalacionTablero(data);
        });
        formInstalacion._listenerAdded = true;
      }
    }

    // ENERGÍA
    if (pageId === 'energia') {
      initAutocomplete('autocompleteEnergia', 'tableroEnergiaId', 'listarTablerosParaEnergia');
      handleProveedorChange('proveedorEnergia', '.detalle-proveedor-energia');
      const formEnergia = document.getElementById('form-energia');
      if (formEnergia && !formEnergia._listenerAdded) {
        formEnergia.addEventListener('submit', e => {
          e.preventDefault();
          const btn = e.submitter; if (btn) btn.disabled = true;
          const id = document.getElementById('tableroEnergiaId')?.value || '';
          const proveedor = document.getElementById('proveedorEnergia')?.value || '';
          const detalleProveedor = document.getElementById('detalleProveedorEnergia')?.value || '';
          const data = {
            id,
            proveedor,
            detalleProveedor: proveedor === 'OTRA' ? detalleProveedor : proveedor,
            observaciones: document.getElementById('obsEnergia')?.value || ''
          };
          if (!id) { M.toast({ html: 'Debe seleccionar un Tablero.' }); if (btn) btn.disabled = false; return; }
          SCRIPT_RUN.withSuccessHandler(res => {
            M.toast({ html: res.success ? 'Energía de ' + id + ' CONECTADA.' : 'Error: ' + res.message, classes: res.success ? 'green darken-1' : 'red darken-3' });
            formEnergia.reset(); M.updateTextFields(); if (btn) btn.disabled = false; showPage('panel');
          }).registrarEnergiaTablero(data);
        });
        formEnergia._listenerAdded = true;
      }
    }

    // CONECTIVIDAD
    if (pageId === 'conectividad') {
      initAutocomplete('autocompleteConectividad', 'tableroConectividadId', 'listarTablerosParaConectividad');
      handleProveedorChange('proveedorConectividad', '.detalle-proveedor-conectividad');
      const formCon = document.getElementById('form-conectividad');
      if (formCon && !formCon._listenerAdded) {
        formCon.addEventListener('submit', e => {
          e.preventDefault();
          const btn = e.submitter; if (btn) btn.disabled = true;
          const id = document.getElementById('tableroConectividadId')?.value || '';
          const proveedor = document.getElementById('proveedorConectividad')?.value || '';
          const detalleProveedor = document.getElementById('detalleProveedorConectividad')?.value || '';
          const tipo = document.getElementById('tipoConectividad')?.value || '';
          const switchName = document.getElementById('filtroSwitch')?.value || '';

          const data = {
            id,
            proveedor,
            detalleProveedor: proveedor === 'OTRA' ? detalleProveedor : proveedor,
            tipo,
            switch: switchName,
            observaciones: document.getElementById('obsConectividad')?.value || ''
          };

          if (!id) { M.toast({ html: 'Debe seleccionar un Tablero.' }); if (btn) btn.disabled = false; return; }

          SCRIPT_RUN.withSuccessHandler(res => {
            M.toast({ html: res.success ? 'Conectividad de ' + id + ' CONECTADA.' : 'Error: ' + res.message, classes: res.success ? 'green darken-1' : 'red darken-3' });
            formCon.reset(); M.updateTextFields(); if (btn) btn.disabled = false; showPage('panel');
          }).registrarConectividadTablero(data);
        });
        formCon._listenerAdded = true;
      }

      // SWITCH
      if (pageId === 'switch') {
        initAutocomplete('autocompleteSwitch', 'tableroSwitchId', 'listarTablerosParaSwitch', true);
        handleProveedorChange('proveedorConectividadSwitch', '.detalle-proveedor-switch');
        const formSwitch = document.getElementById('form-switch');
        if (formSwitch && !formSwitch._listenerAdded) {
          formSwitch.addEventListener('submit', e => {
            e.preventDefault();
            const btn = e.submitter; if (btn) btn.disabled = true;
            const id = document.getElementById('tableroSwitchId')?.value || '';
            const switchName = document.getElementById('inputSwitch')?.value || '';
            const proveedor = document.getElementById('proveedorConectividadSwitch')?.value || '';
            const detalleProveedor = document.getElementById('detalleProveedorSwitch')?.value || '';
            const tipo = document.getElementById('tipoConectividadSwitch')?.value || '';
            const data = {
              id,
              switch: switchName,
              proveedor,
              detalleProveedor: proveedor === 'OTRA' ? detalleProveedor : proveedor,
              tipo,
              observaciones: document.getElementById('obsSwitch')?.value || ''
            };
            if (!id) { M.toast({ html: 'Debe seleccionar un Tablero.' }); if (btn) btn.disabled = false; return; }
            SCRIPT_RUN.withSuccessHandler(res => {
              M.toast({ html: res.success ? 'Switch ' + switchName + ' vinculado.' : 'Error: ' + res.message, classes: res.success ? 'green darken-1' : 'red darken-3' });
              formSwitch.reset(); M.updateTextFields(); if (btn) btn.disabled = false; showPage('panel');
            }).vincularSwitchTablero(data);
          });
          formSwitch._listenerAdded = true;
        }
      }

      // INFORMES
      if (pageId === 'informes') {
        const formFiltros = document.getElementById('form-filtros');
        if (formFiltros && !formFiltros._listenerAdded) {
          formFiltros.addEventListener('submit', e => {
            e.preventDefault();
            const filtros = {
              estado: document.getElementById('filtroEstado')?.value || '',
              energia: document.getElementById('filtroEnergia')?.value || '',
              conectividad: document.getElementById('filtroConectividad')?.value || '',
              direccionContains: document.getElementById('filtroDireccion')?.value || '',
              switchContains: document.getElementById('filtroSwitch')?.value || ''
            }; // 👈 cierre correcto del objeto

            SCRIPT_RUN.withSuccessHandler(generarTablaVistaPrevia).listarTablerosFiltrados(filtros);
          });
          formFiltros._listenerAdded = true;
        }

        const btnDescargar = document.getElementById('btnDescargarInforme');
        if (btnDescargar && !btnDescargar._listenerAdded) {
          btnDescargar.addEventListener('click', () => {
            btnDescargar.disabled = true;
            const filtrosActuales = {
              estado: document.getElementById('filtroEstado')?.value || '',
              energia: document.getElementById('filtroEnergia')?.value || '',
              conectividad: document.getElementById('filtroConectividad')?.value || '',
              direccionContains: document.getElementById('filtroDireccion')?.value || '',
              switchContains: document.getElementById('filtroSwitch')?.value || ''
            };
            SCRIPT_RUN.withSuccessHandler(base64 => {
              try {
                const dataUrl = 'data:application/pdf;base64,' + base64;
                const a = document.createElement('a');
                const ts = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
                a.href = dataUrl;
                a.download = 'Informe_Tableros_' + ts + '.pdf';
                document.body.appendChild(a);
                a.click();
                a.remove();
                M.toast({ html: 'Informe descargado correctamente.', classes: 'green darken-1' });
              } catch (err) {
                console.error(err);
                M.toast({ html: 'Error generando descarga.', classes: 'red darken-3' });
              } finally {
                btnDescargar.disabled = false;
              }
            }).generarInformePDF(filtrosActuales);
          });
          btnDescargar._listenerAdded = true;
        }
      }



      // Limpia informe cuando salís de esa sección
      if (pageId !== 'informes') {
        const cont = document.getElementById('tablaResultadosInforme');
        if (cont) cont.innerHTML = '<p class="center-align grey-text">Presione "Vista Previa" para mostrar los resultados.</p>';
        const btnDesc = document.getElementById('btnDescargarInforme');
        if (btnDesc) btnDesc.disabled = true;
        const countEl = document.getElementById('countDescarga');
        if (countEl) countEl.textContent = 0;
      }

      // Re-inicializar selects
      document.querySelectorAll('select').forEach(el => {
        const instance = M.FormSelect.getInstance(el);
        if (instance) instance.destroy();
        M.FormSelect.init(el);
      });
    } // 👈 este es el cierre correcto de window.showPage


    // --- Función abrir mapa ---
    function abrirMapa() {
      const mapUrl = "https://www.google.com/maps/d/edit?mid=16ZDkCjQQJKOO9w6t7rGCaQ9ibLUH8ps&ll=-34.626221418673786%2C-59.42575859457151&z=17";
      window.open(mapUrl, '_blank');
      M.toast({ html: 'Abriendo Mapa en una nueva pestaña...', classes: 'blue' });
      if (SCRIPT_RUN.registrarAccesoUsuario) {
        SCRIPT_RUN.registrarAccesoUsuario('Página: MAPA');
      }
    }

    // --- Navegación: tarjetas y botones Volver ---
    document.querySelectorAll('.boton-tarjeta[data-page], .btn-volver[data-page]').forEach(el => {
      if (!el) return;
      el.addEventListener('click', () => {
        const pageId = el.getAttribute('data-page');
        if (pageId === 'mapa') abrirMapa();
        else window.showPage(pageId);
      });
    });

    // Inicializar selects al inicio
    document.querySelectorAll('select').forEach(el => M.FormSelect.init(el));
    // Página inicial
    window.showPage('panel');
  }); // 👈 cierre único del document.addEventListener



function generarTablaVistaPrevia(data) {
  const container = document.getElementById('tablaResultadosInforme');
  const btnDescarga = document.getElementById('btnDescargarInforme');
  const countDescarga = document.getElementById('countDescarga');
  container.innerHTML = '';

  if (!data || data.length === 0) {
    container.innerHTML = '<p class="center-align">No se encontraron resultados con los filtros aplicados.</p>';
    btnDescarga.disabled = true;
    countDescarga.textContent = 0;
    return;
  }

  countDescarga.textContent = data.length;
  btnDescarga.disabled = false;

  let html = '<table class="striped highlight responsive-table"><thead><tr><th>ID</th><th>Dirección</th><th>Estado</th><th>Energía</th><th>Conectividad</th><th>Proveedor</th></tr></thead><tbody>';

  data.slice(0, 400).forEach(r => {
    const classEstado = getClassForStatus(r.estado);
    const classEnergia = getClassForStatus(r.energia);
    const classConectividad = getClassForStatus(r.conectividad);

    html += '<tr><td>' + r.id + '</td><td>' + r.direccion + '</td><td class="' + classEstado + '">' + r.estado + '</td><td class="' + classEnergia + '">' + r.energia + '</td><td class="' + classConectividad + '">' + r.conectividad + '</td><td>' + r.proveedor + '</td></tr>';
  });

  html += '</tbody></table>';

  if (data.length > 400) {
    html += '<p class="center-align" style="font-style: italic; margin-top: 15px;">Mostrando las primeras filas. El PDF incluirá las ' + data.length + ' filas.</p>';
  }

  container.innerHTML = html;
}
function getClassForStatus(status) {
  status = status ? status.toUpperCase() : '';
  switch (status) {
    case 'PEDIDO': return 'amber lighten-4';
    case 'INSTALADO': return 'red lighten-4';
    case 'CONECTADA': return 'green lighten-4';
    case 'PENDIENTE': return 'deep-orange lighten-4';
    default: return '';
  }
}