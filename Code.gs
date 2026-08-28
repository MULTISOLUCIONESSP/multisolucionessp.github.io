/**
 * CODE.GS — Backend de Google Apps Script para "La Cocina de Mamá"
 * -------------------------------------------------------------------
 * Estructura de la hoja "Productos" (fila 1 = encabezados):
 *   A: id_producto
 *   B: categoria            (economico / clasico / gourmet)
 *   C: nombre
 *   D: descripcion
 *   E: precio               (precio normal, el de "Menú Clásico")
 *   F: disponible          ("si" / "no")
 *   G: sotck_actual        (stock actual)
 *   H: imagen_url          (foto principal del plato)
 *   I: variante            (opcional — nombres separados por coma,
 *                            ej: "Salsa Bolognesa,Salsa Bechamel")
 *   J: imagen_variante_url (opcional — URLs en el MISMO orden que
 *                            "variante", separadas por coma)
 *   K: es_plato_del_dia    (🔧 "si" / "no" - reemplaza a "grupo_principal".
 *                            TODOS los productos son de "Menú Clásico" por
 *                            defecto - tildando esto, el MISMO producto
 *                            (sin duplicar la fila) también aparece del
 *                            lado de "Plato del Día", en su misma categoría.)
 *   L: precio_plato_del_dia (🆕 opcional - precio distinto para cuando se
 *                            muestra como Plato del Día. Si se deja vacío,
 *                            usa el precio normal de la columna E.)
  *   M: variante_foto_principal (🆕 opcional - a cuál de las variantes de
 *                            la columna I corresponde la FOTO PRINCIPAL
 *                            del plato. Si el cliente suma porciones sin
 *                            abrir el desplegable de variantes, esto es
 *                            lo que se usa - así nunca queda "sin
 *                            variante" en la comanda para un plato que
 *                            sí tiene variantes.)
 *   N: es_premio            (🆕 "si" / "no" - ¿también se puede canjear
 *                            con puntos de fidelidad?)
 *   O: puntos_requeridos    (🆕 cuántos puntos hacen falta para canjearlo)
 * -------------------------------------------------------------------
 */

const NOMBRE_HOJA_PRODUCTOS = "Productos";
const ID_HOJA_CALCULO = "1Sphs8hsSJ9_1prgfbjd0rjCrNb2C8xarrNrY7nOXjqQ";

function obtenerHojaCalculo() {
  return SpreadsheetApp.openById(ID_HOJA_CALCULO);
}

const ADMIN_USUARIO = "admin";
const ADMIN_CLAVE = "123456";

// 🆕 Carpeta de Drive donde se guardan las fotos subidas desde el panel de
// admin - se crea sola la primera vez que hace falta.
const NOMBRE_CARPETA_FOTOS = "Fotos - La Cocina de Mamá";

function doGet(e) {
  if (e && e.parameter && e.parameter.panel === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Panel de Administración')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 🆕 Solo la vista del cliente permite mostrarse dentro de un iframe -
  // necesario para la versión "app" (TWA) que envuelve esta página
  // desde un dominio propio (GitHub Pages). El panel de admin NO se
  // toca, sigue con la protección normal.
  return HtmlService.createHtmlOutputFromFile('menus')
    .setTitle('La Cocina de Mamá')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validarAdmin(usuario, clave) {
  const usuarioOk = String(usuario).trim().toLowerCase() === ADMIN_USUARIO.toLowerCase();
  const claveOk = String(clave).trim() === ADMIN_CLAVE;
  return usuarioOk && claveOk;
}

function obtenerUrlApp() {
  return ScriptApp.getService().getUrl();
}

// ==========================================================================
// 🆕 SUBIDA REAL DE IMÁGENES A DRIVE - mismo patrón exacto que ya usaba
// Mercado Artesanal ("subirImagenDrive"): recibe el archivo ya comprimido
// en base64 desde el navegador, lo decodifica, lo guarda en una carpeta
// propia de Drive, lo deja visible para "cualquiera con el enlace"
// (imprescindible para poder mostrarlo después en la app), y devuelve
// solo el ID del archivo (no la URL completa) - el resto del sistema ya
// sabe convertir un ID en la URL de imagen que necesite.
// ==========================================================================
// 🆕 AUTORIZACIÓN DE DRIVE - correr esta función UNA SOLA VEZ, directo
// desde el editor de Apps Script (no desde la app publicada), para que
// aparezca el cartel pidiendo permiso de Drive y lo aceptes. Una vez
// aceptado, la app publicada ya puede subir fotos sin problema - se
// puede borrar esta función después si querés, no la usa nadie más.
// 🔧 FIX: antes solo LEÍA de Drive (getRootFolder) - eso solo pedía el
// permiso de lectura, no el de escritura que hace falta para crear
// carpetas y archivos (createFolder/createFile). Ahora también CREA y
// borra una carpeta de prueba, para forzar que Google pida el permiso
// completo de una sola vez.
function autorizarAccesoADrive() {
  const carpeta = DriveApp.getRootFolder();
  Logger.log("✅ Permiso de lectura OK - carpeta raíz: " + carpeta.getName());

  const carpetaPrueba = DriveApp.createFolder("__prueba_permiso_borrar__");
  Logger.log("✅ Permiso de escritura OK - carpeta de prueba creada.");
  carpetaPrueba.setTrashed(true); // la manda a la papelera, no queda basura
  Logger.log("✅ Listo - los 2 permisos de Drive quedaron concedidos.");
}

function obtenerOCrearCarpetaFotos() {
  const carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_FOTOS);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(NOMBRE_CARPETA_FOTOS);
}

function subirImagenDrive(paqueteArchivo) {
  try {
    if (!paqueteArchivo || !paqueteArchivo.base64) return null;

    var base64Crudo = paqueteArchivo.base64;
    if (base64Crudo.indexOf(",") !== -1) {
      base64Crudo = base64Crudo.split(",")[1]; // saca el prefijo "data:image/jpeg;base64,"
    }

    const bytes = Utilities.base64Decode(base64Crudo);
    const nombreLimpio = "plato_" + new Date().getTime() + "_" + String(paqueteArchivo.name || "foto").replace(/[^a-zA-Z0-9.]/g, "_");
    const blob = Utilities.newBlob(bytes, paqueteArchivo.type || "image/jpeg", nombreLimpio);

    const carpeta = obtenerOCrearCarpetaFotos();
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return archivo.getId();
  } catch (err) {
    Logger.log("Fallo al subir imagen a Drive: " + err.message);
    return null;
  }
}

// 🆕 Sube la foto principal (si viene una nueva) y las de cada variante
// que también haya cambiado - las que no cambiaron (edición de un
// producto ya cargado) se dejan con su link viejo tal cual estaba.
// Devuelve { imagenUrl, variante, imagenVarianteUrl } listos para
// guardar en la hoja, en el mismo formato de siempre.
// 🔧 FIX RAÍZ: antes, si subirImagenDrive() fallaba (por el motivo que
// sea - permisos, tamaño, lo que sea), esto seguía de largo sin avisar
// nada - el producto se guardaba bien con el resto de los datos, pero la
// columna de la foto quedaba vacía, sin ningún error visible. Ahora, si
// se mandó una foto NUEVA y la subida falla, se corta acá con un error
// claro, en vez de guardar el producto a medias en silencio.
function procesarImagenesProducto(datos) {
  var imagenUrlFinal = datos.imagenUrlExistente || "";
  if (datos.imagenArchivo && datos.imagenArchivo.base64) {
    var idNuevo = subirImagenDrive(datos.imagenArchivo);
    if (!idNuevo) {
      throw new Error("La foto principal no se pudo subir a Drive. Revisá 'Ejecuciones' en el editor de Apps Script para ver el motivo exacto.");
    }
    imagenUrlFinal = "https://drive.google.com/file/d/" + idNuevo + "/view";
  }

  var nombresVariante = [];
  var urlsVariante = [];
  (datos.variantes || []).forEach(function(v) {
    if (!v.nombre) return;
    nombresVariante.push(v.nombre);
    var urlVarianteFinal = v.urlExistente || "";
    if (v.archivo && v.archivo.base64) {
      var idVariante = subirImagenDrive(v.archivo);
      if (!idVariante) {
        throw new Error("La foto de la variante '" + v.nombre + "' no se pudo subir a Drive. Revisá 'Ejecuciones' en el editor de Apps Script para ver el motivo exacto.");
      }
      urlVarianteFinal = "https://drive.google.com/file/d/" + idVariante + "/view";
    }
    urlsVariante.push(urlVarianteFinal);
  });

  return {
    imagenUrl: imagenUrlFinal,
    variante: nombresVariante.join(','),
    imagenVarianteUrl: urlsVariante.join(',')
  };
}

// Agrega una fila nueva a la hoja de Productos desde el panel de administración.
function agregarProducto(datos) {
  try {
    const hoja = obtenerHojaCalculo().getSheetByName(NOMBRE_HOJA_PRODUCTOS);
    if (!hoja) return { ok: false, error: 'No se encontró la hoja "' + NOMBRE_HOJA_PRODUCTOS + '".' };

    var imagenes = procesarImagenesProducto(datos);

    hoja.appendRow([
      datos.idProducto,
      datos.categoria,
      datos.nombre,
      datos.descripcion,
      datos.precio,
      datos.disponible,
      datos.stockActual,
      imagenes.imagenUrl,
      imagenes.variante,
      imagenes.imagenVarianteUrl,
      datos.esPlatoDelDia || 'no', // 🔧 K: ¿también aparece como Plato del Día?
      datos.precioPlatoDelDia || '', // 🆕 L: precio distinto para Plato del Día (opcional)
      datos.varianteFotoPrincipal || '', // 🆕 M: a qué variante corresponde la foto principal
      datos.esPremio || 'no', // 🆕 N: ¿es canjeable con puntos?
      datos.puntosRequeridos || '' // 🆕 O: cuántos puntos hacen falta
    ]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: "Error al guardar: " + err.message };
  }
}

// Lee la hoja de Productos y arma el array que usa el frontend (menús).
function obtenerProductos() {
  const hoja = obtenerHojaCalculo().getSheetByName(NOMBRE_HOJA_PRODUCTOS);
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const filas = datos.slice(1);

  const productos = [];
  filas.forEach(fila => {
    const [idProducto, categoria, nombre, descripcion, precio, disponible, stockActual, imagenUrl, variante, imagenVarianteUrl, esPlatoDelDia, precioPlatoDelDia, varianteFotoPrincipal, esPremio, puntosRequeridos] = fila;
    if (!nombre || String(nombre).trim() === '') return;
    if (String(disponible).trim().toLowerCase() !== 'si') return;

    productos.push({
      id: String(idProducto),
      categoria: String(categoria).trim().toLowerCase(),
      nombre: String(nombre).trim(),
      descripcion: descripcion ? String(descripcion).trim() : '',
      precio: Number(precio) || 0,
      stock: Number(stockActual) || 0,
      imagen_url: imagenUrl ? String(imagenUrl).trim() : '',
      variante: variante ? String(variante).trim() : '',
      imagen_variante_url: imagenVarianteUrl ? String(imagenVarianteUrl).trim() : '',
      // 🔧 Reemplaza a "grupo_principal" - TODO producto es de Menú
      // Clásico por defecto; este flag indica si TAMBIÉN se muestra del
      // lado de Plato del Día (misma fila, sin duplicar nada).
      es_plato_del_dia: String(esPlatoDelDia).trim().toLowerCase() === 'si',
      precio_plato_del_dia: precioPlatoDelDia ? Number(precioPlatoDelDia) : null,
      // 🆕 A qué variante corresponde la foto principal (si el plato
      // tiene variantes) - así sumar sin abrir el desplegable no queda
      // "sin variante" en la comanda.
      variante_foto_principal: varianteFotoPrincipal ? String(varianteFotoPrincipal).trim() : '',
      // 🆕 Si es canjeable con puntos, y cuántos hacen falta.
      es_premio: String(esPremio).trim().toLowerCase() === 'si',
      puntos_requeridos: puntosRequeridos ? Number(puntosRequeridos) : 0
    });
  });

  return productos;
}

// Lista TODOS los productos (disponibles o no) para el panel de administración.
function obtenerProductosAdmin() {
  const hoja = obtenerHojaCalculo().getSheetByName(NOMBRE_HOJA_PRODUCTOS);
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const productos = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const [idProducto, categoria, nombre, descripcion, precio, disponible, stockActual, imagenUrl, variante, imagenVarianteUrl, esPlatoDelDia, precioPlatoDelDia, varianteFotoPrincipal, esPremio, puntosRequeridos] = fila;
    if (!nombre || String(nombre).trim() === '') continue;

    productos.push({
      fila: i + 1,
      idProducto: String(idProducto),
      categoria: categoria,
      nombre: nombre,
      descripcion: descripcion,
      precio: precio,
      disponible: disponible,
      stockActual: stockActual,
      imagenUrl: imagenUrl,
      variante: variante || '',
      imagenVarianteUrl: imagenVarianteUrl || '',
      esPlatoDelDia: esPlatoDelDia || 'no',
      precioPlatoDelDia: precioPlatoDelDia || '',
      varianteFotoPrincipal: varianteFotoPrincipal || '',
      esPremio: esPremio || 'no',
      puntosRequeridos: puntosRequeridos || ''
    });
  }

  return productos;
}

// Actualiza una fila existente (identificada por "fila") con los datos nuevos.
function actualizarProducto(datos) {
  try {
    const hoja = obtenerHojaCalculo().getSheetByName(NOMBRE_HOJA_PRODUCTOS);
    if (!hoja) return { ok: false, error: 'No se encontró la hoja "' + NOMBRE_HOJA_PRODUCTOS + '".' };
    if (!datos.fila) return { ok: false, error: 'Falta indicar qué producto actualizar.' };

    var imagenes = procesarImagenesProducto(datos);

    hoja.getRange(datos.fila, 1, 1, 15).setValues([[
      datos.idProducto,
      datos.categoria,
      datos.nombre,
      datos.descripcion,
      datos.precio,
      datos.disponible,
      datos.stockActual,
      imagenes.imagenUrl,
      imagenes.variante,
      imagenes.imagenVarianteUrl,
      datos.esPlatoDelDia || 'no',
      datos.precioPlatoDelDia || '',
      datos.varianteFotoPrincipal || '',
      datos.esPremio || 'no',
      datos.puntosRequeridos || ''
    ]]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: "Error al actualizar: " + err.message };
  }
}

/**
 * ---------------------------------------------------------------
 * LOGIN Y REGISTRO — borrador mínimo, reemplazar por la lógica real
 * cuando esté lista la hoja "Usuarios".
 * ---------------------------------------------------------------
 */
// ==========================================================================
// 🆕 TARIFAS DE ENVÍO POR ZONA (palabra clave) - hoja "Tarifas_Envios",
// que el vendedor puede editar libremente desde el Sheets (agregar zonas,
// cambiar precios) sin tocar nunca el código. El frontend la lee entera
// una sola vez y escanea la dirección tipiada buscando coincidencias.
// ==========================================================================
const NOMBRE_HOJA_TARIFAS = "Tarifas_Envios";

function obtenerOCrearHojaTarifas() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_TARIFAS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_TARIFAS);
    hoja.appendRow(["id_zona", "nombre_zona", "palabra_clave", "costo_envio"]);
    // ⚠️ Filas de EJEMPLO - editalas/agregales más desde el Sheets
    // directamente, sin tocar código. "palabra_clave" se busca en
    // minúsculas, sin importar mayúsculas, dentro de lo que tipeó el
    // cliente en la dirección.
    hoja.appendRow([1, "La Reja Centro", "reja", 900]);
    hoja.appendRow([2, "Moreno Centro", "moreno centro", 700]);
    hoja.appendRow([3, "Moreno (Gral.)", "moreno", 800]);
    hoja.appendRow([4, "Mitre", "mitre", 750]);
  }
  return hoja;
}

// Devuelve la tabla completa de tarifas, para que el frontend escanee la
// dirección del lado del cliente (instantáneo, sin ida y vuelta al
// servidor por cada letra que tipea).
function obtenerTarifasEnvios() {
  try {
    const hoja = obtenerOCrearHojaTarifas();
    const filas = hoja.getDataRange().getValues();
    const tarifas = [];
    for (let i = 1; i < filas.length; i++) {
      if (!filas[i][1]) continue; // saltea filas vacías
      tarifas.push({
        idZona: filas[i][0],
        nombreZona: String(filas[i][1]).trim(),
        palabraClave: String(filas[i][2]).trim().toLowerCase(),
        costoEnvio: Number(filas[i][3]) || 0
      });
    }
    return tarifas;
  } catch (err) {
    return [];
  }
}

// ==========================================================================
// 🆕 PEDIDOS - hoja "Pedidos". Estructura ampliada respecto a la
// original (agrega teléfono, fecha/hora separadas para poder ordenar y
// agrupar bien, tipo de destino, y una versión JSON estructurada del
// detalle - así el panel puede sumar cantidades de forma confiable, sin
// tener que "leer" el texto armado para humanos).
//   A: id_pedido
//   B: fecha_hora        (cuándo se cargó el pedido)
//   C: id_usuario
//   D: nombre_cliente
//   E: telefono          (🆕 para el botón de WhatsApp directo en el panel)
//   F: tipo_destino      (🆕 retiro / ruta / individual)
//   G: empresa_destino   (nombre de empresa, o dirección, según el tipo)
//   H: fecha_entrega     (🆕 separada de la hora, para poder agrupar por día)
//   I: hora_entrega      (🆕 separada de la fecha, para poder ordenar por horario)
//   J: detalle_items     (texto legible, para mostrar en la tarjeta)
//   K: detalle_items_json (🆕 mismo detalle pero estructurado, para que el
//                           "Resumen de cocina" pueda sumar cantidades)
//   L: total
//   M: estado_pago
//   N: estado_pedido
// ==========================================================================
const NOMBRE_HOJA_PEDIDOS = "Pedidos";

const ENCABEZADOS_PEDIDOS = [
  "id_pedido", "fecha_hora", "id_usuario", "nombre_cliente", "telefono",
  "tipo_destino", "empresa_destino", "fecha_entrega", "hora_entrega",
  "detalle_items", "detalle_items_json", "total", "estado_pago", "estado_pedido",
  "puntos_acreditados" // 🆕 "si"/"no" - evita sumar puntos 2 veces, sin
                        // importar en qué orden se toquen "Entregado" y
                        // "Marcar pagado"
];

function obtenerOCrearHojaPedidos() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_PEDIDOS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_PEDIDOS);
    hoja.appendRow(ENCABEZADOS_PEDIDOS);
  } else {
    // 🔧 FIX RAÍZ: si la hoja YA existía de antes (con la estructura
    // vieja de 10 columnas), nunca se actualizaban los títulos - los
    // datos nuevos (14 columnas) se seguían guardando bien, pero las
    // columnas K a N quedaban sin ningún título arriba. Esto repara
    // sola la fila de encabezados si detecta que está desactualizada -
    // NUNCA toca ni borra ninguna fila de datos, solo la fila 1.
    const encabezadosActuales = hoja.getRange(1, 1, 1, ENCABEZADOS_PEDIDOS.length).getValues()[0];
    const desactualizada = ENCABEZADOS_PEDIDOS.some((titulo, i) => encabezadosActuales[i] !== titulo);
    if (desactualizada) {
      hoja.getRange(1, 1, 1, ENCABEZADOS_PEDIDOS.length).setValues([ENCABEZADOS_PEDIDOS]);
    }
  }
  return hoja;
}

function generarIdPedido() {
  return "P-" + new Date().getTime().toString(36).toUpperCase();
}

function registrarPedido(datos) {
  try {
    // 🆕 Si el cliente tiene un recargo pendiente (por haber cancelado
    // tarde un pedido anterior), se le suma acá, y se descuenta de su
    // cuenta - solo se cobra una vez.
    let recargoAplicado = 0;
    if (datos.idUsuario && datos.idUsuario !== 'INVITADO') {
      const hojaUsuarios = obtenerOCrearHojaUsuarios();
      const filaUsuario = encontrarFilaUsuario(hojaUsuarios, datos.idUsuario);
      if (filaUsuario) {
        const recargoPendiente = Number(hojaUsuarios.getRange(filaUsuario, 14).getValue()) || 0;
        if (recargoPendiente > 0) {
          recargoAplicado = recargoPendiente;
          hojaUsuarios.getRange(filaUsuario, 14).setValue(0);
        }
      }
    }

    // 🆕 REGLA DE NEGOCIO: los premios canjeados viajan JUNTOS con el
    // pedido (no se registran aparte) - y solo se pueden incluir si el
    // pedido tiene al menos un plato principal de verdad. Esto también
    // evita el bug de antes, donde un canje "solo" generaba SU PROPIO
    // pedido y de yapa sumaba puntos por ese "pedido" fantasma.
    const premios = datos.premiosCanjeados || [];
    if (premios.length > 0) {
      const hayPlatoPrincipal = (datos.itemsEstructurados || JSON.parse(datos.detalleItemsJson || "[]"))
        .some(item => item.tipo === 'plato');
      if (!hayPlatoPrincipal) {
        return { ok: false, error: 'Para canjear un premio, tu pedido necesita al menos un plato principal.' };
      }

      // Valida que le alcancen los puntos para TODOS los premios juntos,
      // y los descuenta recién acá, todo o nada.
      const hojaUsuarios = obtenerOCrearHojaUsuarios();
      const filaUsuario = encontrarFilaUsuario(hojaUsuarios, datos.idUsuario);
      if (!filaUsuario) return { ok: false, error: 'No se encontró tu usuario.' };

      const puntosActuales = Number(hojaUsuarios.getRange(filaUsuario, 11).getValue()) || 0;
      const puntosNecesarios = premios.reduce((acc, p) => acc + (Number(p.puntosRequeridos) || 0), 0);
      if (puntosActuales < puntosNecesarios) {
        return { ok: false, error: 'No te alcanzan los puntos para ese/esos premio(s).' };
      }
      hojaUsuarios.getRange(filaUsuario, 11).setValue(puntosActuales - puntosNecesarios);
    }

    const hoja = obtenerOCrearHojaPedidos();
    const idPedido = generarIdPedido();

    const estadoPago = (datos.formaPago === "transferencia") ? "Pendiente de confirmar" : "A cobrar al entregar";

    // 🔧 FIX RAÍZ: "fecha_entrega" (columna H, ej: "2026-08-27") y
    // "hora_entrega" (columna I, ej: "5:55") son TEXTO a propósito -
    // pero Sheets detecta ese patrón y las convierte solas en una fecha/
    // hora de verdad (mismo problema que ya tuvimos antes con las
    // coordenadas del mapa). Eso rompe el agrupado por día en el panel,
    // que espera el texto exacto. Por eso, antes de escribir, se fuerza
    // el formato de esas celdas a texto plano ("@") - incluye el
    // teléfono (columna E), que Sheets también detecta como número
    // largo y lo convierte solo, rompiendo el .replace() del panel.
    const filaDestino = hoja.getLastRow() + 1;
    hoja.getRange(filaDestino, 5).setNumberFormat("@"); // E: telefono
    hoja.getRange(filaDestino, 8).setNumberFormat("@"); // H: fecha_entrega
    hoja.getRange(filaDestino, 9).setNumberFormat("@"); // I: hora_entrega

    const totalConRecargo = (datos.total || 0) + recargoAplicado;
    const detalleConRecargo = datos.detalleItems + (recargoAplicado > 0 ? ' | ⚠️ Recargo por cancelación tardía anterior: $' + recargoAplicado : '');

    hoja.getRange(filaDestino, 1, 1, 15).setValues([[
      idPedido,
      new Date(),
      datos.idUsuario || "INVITADO",
      datos.nombreCliente || "",
      datos.telefono || "",
      datos.tipoDestino || "",
      datos.empresaDestino || "",
      datos.fechaEntrega || "",
      datos.horaEntrega || "",
      detalleConRecargo || "",
      datos.detalleItemsJson || "[]",
      totalConRecargo,
      estadoPago,
      "Pendiente",
      "no"
    ]]);

    return { ok: true, idPedido: idPedido, recargoAplicado: recargoAplicado };
  } catch (err) {
    Logger.log("Error al registrar pedido: " + err.message);
    return { ok: false, error: err.message };
  }
}

// 🆕 Trae TODOS los pedidos para el panel de admin, ya armados en el
// formato que necesita la pantalla de "Pedidos" (agrupados por día y
// tipo de destino en el frontend, esto solo devuelve la lista plana).
// 🆕 Si "fechaEntrega"/"horaEntrega" ya quedaron guardadas como fecha/hora
// real de Sheets (pedidos cargados ANTES de este arreglo), esto las
// vuelve a convertir al texto esperado ("2026-08-27", "17:30") - así el
// panel agrupa bien esos pedidos viejos también, sin tener que
// corregirlos a mano en la planilla.
function normalizarFechaTexto(valor) {
  if (!valor) return "";
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(valor);
}
function normalizarHoraTexto(valor) {
  if (!valor) return "";
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(valor);
}

function obtenerPedidosAdmin() {
  const hoja = obtenerOCrearHojaPedidos();
  const datos = hoja.getDataRange().getValues();
  const pedidos = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const [idPedido, fechaHora, idUsuario, nombreCliente, telefono, tipoDestino,
           empresaDestino, fechaEntrega, horaEntrega, detalleItems, detalleItemsJson,
           total, estadoPago, estadoPedido] = fila;
    if (!idPedido) continue;

    let itemsEstructurados = [];
    try { itemsEstructurados = JSON.parse(detalleItemsJson || "[]"); } catch (e) { itemsEstructurados = []; }

    pedidos.push({
      fila: i + 1,
      idPedido: String(idPedido),
      idUsuario: idUsuario || "",
      nombreCliente: nombreCliente || "",
      telefono: telefono ? String(telefono) : "",
      tipoDestino: tipoDestino || "",
      empresaDestino: empresaDestino || "",
      fechaEntrega: normalizarFechaTexto(fechaEntrega),
      horaEntrega: normalizarHoraTexto(horaEntrega),
      detalleItems: detalleItems || "",
      itemsEstructurados: itemsEstructurados,
      total: Number(total) || 0,
      estadoPago: estadoPago || "",
      estadoPedido: estadoPedido || "Pendiente"
    });
  }

  return pedidos;
}

// 🆕 Cambia el estado de un pedido (Pendiente / En preparación /
// Entregado), o el estado de pago (Pendiente de confirmar / Pagado).
function actualizarEstadoPedido(fila, campo, valorNuevo) {
  try {
    const hoja = obtenerOCrearHojaPedidos();

    const columna = (campo === 'estadoPago') ? 13 : 14; // M=estado_pago, N=estado_pedido
    hoja.getRange(fila, columna).setValue(valorNuevo);

    // 🆕 REGLA DE NEGOCIO: los puntos se acreditan solo cuando el PAGO
    // está confirmado Y el pedido está Entregado - no alcanza con uno
    // solo. Como "Entregado" y "Marcar pagado" pueden tocarse en
    // cualquier orden, esto se fija leyendo el estado ACTUAL de ambas
    // columnas después de aplicar el cambio, en vez de asumir cuál se
    // tocó primero. La columna O (puntos_acreditados) evita duplicar,
    // sin importar cuántas veces se cumpla la condición de nuevo.
    const estadoPagoActual = hoja.getRange(fila, 13).getValue(); // M: estado_pago
    const estadoPedidoActual = hoja.getRange(fila, 14).getValue(); // N: estado_pedido
    const yaAcreditados = hoja.getRange(fila, 15).getValue(); // O: puntos_acreditados

    // "A cobrar al entregar" (efectivo) se considera pago confirmado en
    // el momento mismo de la entrega - "Pendiente de confirmar"
    // (transferencia) todavía no.
    const pagoConfirmado = estadoPagoActual === 'Pagado' || estadoPagoActual === 'A cobrar al entregar';

    if (estadoPedidoActual === 'Entregado' && pagoConfirmado && yaAcreditados !== 'si') {
      const idUsuarioDelPedido = hoja.getRange(fila, 3).getValue(); // C: id_usuario
      sumarPuntosPorPedidoEntregado(idUsuarioDelPedido);
      hoja.getRange(fila, 15).setValue('si');
    }

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ==========================================================================
// 🆕 USUARIOS - login y registro reales (antes eran solo bocetos que
// siempre devolvían "no"). Estructura de la hoja "Usuarios":
//   A: id_usuario
//   B: nombre
//   C: email
//   D: clave
//   E: telefono
//   F: empresa_zona
//   G: provincia
//   H: zona
//   I: localidad
//   J: direccion
//   K: puntos       (arranca en 0, suma 250 cada vez que un pedido de
//                     ese usuario se marca "Entregado" desde el panel)
//   L: fecha_registro
// ==========================================================================
const NOMBRE_HOJA_USUARIOS = "Usuarios";
const PUNTOS_POR_PEDIDO_ENTREGADO = 250;

const ENCABEZADOS_USUARIOS = [
  "id_usuario", "nombre", "email", "clave", "telefono", "empresa_zona",
  "provincia", "zona", "localidad", "direccion", "puntos", "fecha_registro", "favoritos",
  "recargo_pendiente" // 🆕 se suma al PRÓXIMO pedido, si canceló tarde uno anterior
];

function obtenerOCrearHojaUsuarios() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_USUARIOS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_USUARIOS);
    hoja.appendRow(ENCABEZADOS_USUARIOS);
  } else {
    // 🔧 Tu hoja real tenía menos columnas (id_usuario, nombre,
    // empresa_o_zona, puntos_fidelidad, usuario_verificado) - como el
    // registro nunca funcionó de verdad hasta ahora, no hay datos reales
    // guardados todavía, así que esto actualiza los encabezados a la
    // estructura completa sin perder nada.
    const encabezadosActuales = hoja.getRange(1, 1, 1, ENCABEZADOS_USUARIOS.length).getValues()[0];
    const desactualizada = ENCABEZADOS_USUARIOS.some((titulo, i) => encabezadosActuales[i] !== titulo);
    if (desactualizada) {
      hoja.getRange(1, 1, 1, ENCABEZADOS_USUARIOS.length).setValues([ENCABEZADOS_USUARIOS]);
    }
  }
  return hoja;
}

// Arma el objeto de usuario que se manda al frontend - SIN la contraseña,
// nunca hace falta del otro lado y no tiene sentido exponerla.
function armarUsuarioParaFrontend(fila, numeroFila) {
  return {
    fila: numeroFila,
    idUsuario: fila[0],
    nombre: fila[1],
    email: fila[2],
    telefono: fila[4],
    empresaZona: fila[5],
    provincia: fila[6],
    zona: fila[7],
    localidad: fila[8],
    direccion: fila[9],
    puntos: Number(fila[10]) || 0,
    // 🆕 IDs de productos favoritos, separados por coma.
    favoritos: fila[12] ? String(fila[12]).split(',').map(s => s.trim()).filter(Boolean) : [],
    recargoPendiente: Number(fila[13]) || 0
  };
}

function registrarUsuario(datos) {
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const filas = hoja.getDataRange().getValues();

    // No deja registrar 2 veces el mismo correo.
    const emailNuevo = String(datos.email).trim().toLowerCase();
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][2]).trim().toLowerCase() === emailNuevo) {
        return { ok: false, error: 'Ese correo ya tiene una cuenta - iniciá sesión en vez de registrarte.' };
      }
    }

    // El teléfono y la clave se guardan forzados como texto (mismo
    // motivo que en Pedidos - si son puramente numéricos, Sheets los
    // detecta como número y los convierte solo, lo que puede romper
    // contraseñas con ceros a la izquierda, por ejemplo).
    const filaDestino = hoja.getLastRow() + 1;
    hoja.getRange(filaDestino, 4).setNumberFormat("@"); // D: clave
    hoja.getRange(filaDestino, 5).setNumberFormat("@"); // E: telefono

    const nuevaFila = [
      datos.idUsuario,
      datos.nombre,
      datos.email,
      String(datos.clave).trim(),
      datos.telefono,
      datos.empresaZona || '',
      datos.provincia || '',
      datos.zona || '',
      datos.localidad || '',
      datos.direccion || '',
      0, // puntos arranca en 0
      new Date(),
      '', // favoritos arranca vacío
      0 // recargo pendiente arranca en 0
    ];
    hoja.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    return { ok: true, usuario: armarUsuarioParaFrontend(nuevaFila, filaDestino) };
  } catch (err) {
    return { ok: false, error: "Error al registrar: " + err.message };
  }
}

function validarLogin(email, password) {
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const filas = hoja.getDataRange().getValues();
    const emailBuscado = String(email).trim().toLowerCase();

    for (let i = 1; i < filas.length; i++) {
      const fila = filas[i];
      if (String(fila[2]).trim().toLowerCase() === emailBuscado && String(fila[3]).trim() === String(password).trim()) {
        return armarUsuarioParaFrontend(fila, i + 1);
      }
    }
    return null; // no encontrado, o contraseña incorrecta
  } catch (err) {
    return null;
  }
}

// 🆕 Suma PUNTOS_POR_PEDIDO_ENTREGADO al usuario dueño de un pedido, la
// primera vez que ese pedido se marca como "Entregado" - se llama desde
// actualizarEstadoPedido() de más abajo. Los invitados (sin cuenta) no
// tienen fila en Usuarios, así que no suman nada - no rompe si no los
// encuentra.
function sumarPuntosPorPedidoEntregado(idUsuario) {
  if (!idUsuario || idUsuario === 'INVITADO') return;
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const filas = hoja.getDataRange().getValues();
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) === String(idUsuario)) {
        const puntosActuales = Number(filas[i][10]) || 0;
        hoja.getRange(i + 1, 11).setValue(puntosActuales + PUNTOS_POR_PEDIDO_ENTREGADO);
        return;
      }
    }
  } catch (err) {
    Logger.log("Error al sumar puntos: " + err.message);
  }
}

// 🆕 Trae TODOS los usuarios registrados, para la pestaña "Clientes" del panel.
function obtenerUsuariosAdmin() {
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const filas = hoja.getDataRange().getValues();
    const usuarios = [];
    for (let i = 1; i < filas.length; i++) {
      if (!filas[i][0]) continue; // saltea filas vacías
      usuarios.push(armarUsuarioParaFrontend(filas[i], i + 1));
    }
    return usuarios;
  } catch (err) {
    return [];
  }
}

// ==========================================================================
// 🆕 PANEL DEL CLIENTE - perfil editable, pedidos propios, favoritos,
// y canje de premios con puntos.
// ==========================================================================

// Encuentra la fila de un usuario por su ID - función chica compartida
// por varias de las de abajo, para no repetir el mismo bucle.
function encontrarFilaUsuario(hoja, idUsuario) {
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(idUsuario)) return i + 1;
  }
  return null;
}

// 🆕 Actualiza los datos editables del perfil (nombre, teléfono, empresa/
// zona, dirección) - el email y la contraseña NO se tocan acá.
function actualizarPerfilUsuario(datos) {
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const fila = encontrarFilaUsuario(hoja, datos.idUsuario);
    if (!fila) return { ok: false, error: 'No se encontró tu usuario.' };

    hoja.getRange(fila, 5).setNumberFormat("@"); // E: telefono, forzado texto
    hoja.getRange(fila, 2).setValue(datos.nombre || '');
    hoja.getRange(fila, 5).setValue(datos.telefono || '');
    hoja.getRange(fila, 6).setValue(datos.empresaZona || '');
    hoja.getRange(fila, 10).setValue(datos.direccion || '');

    const filaActualizada = hoja.getRange(fila, 1, 1, ENCABEZADOS_USUARIOS.length).getValues()[0];
    return { ok: true, usuario: armarUsuarioParaFrontend(filaActualizada, fila) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// 🆕 Trae los pedidos de UN usuario puntual, para "Mis Pedidos" - mismo
// formato que obtenerPedidosAdmin(), pero filtrado.
function obtenerPedidosDeUsuario(idUsuario) {
  try {
    const todos = obtenerPedidosAdmin();
    return todos.filter(p => String(p.idUsuario) === String(idUsuario));
  } catch (err) {
    return [];
  }
}

// 🆕 Trae los productos marcados como premio, ordenados de MENOS a MÁS
// puntos - para que el cliente vea primero lo que está más cerca de
// poder canjear.
function obtenerPremiosDisponibles() {
  try {
    // 🔧 Los IDs de Bebidas y Postres son independientes entre sí (y de
    // Productos) - podrían repetirse (ej: "1" en las 3 hojas). Se
    // prefijan acá para que nunca se confundan al sacarlos del carrito.
    const platos = obtenerProductos()
      .filter(p => p.es_premio)
      .map(p => ({ id: 'plato_' + p.id, nombre: p.nombre, imagen_url: p.imagen_url, puntos_requeridos: p.puntos_requeridos, tipo: 'plato' }));

    const bebidas = obtenerBebidasAdmin()
      .filter(b => b.disponible && b.esPremio)
      .map(b => ({ id: 'bebida_' + b.id, nombre: b.nombre + ' (' + b.tamano + ')', imagen_url: b.imagenUrl, puntos_requeridos: b.puntosRequeridos, tipo: 'bebida' }));

    const postres = obtenerPostresAdmin()
      .filter(p => p.disponible && p.esPremio)
      .map(p => ({ id: 'postre_' + p.id, nombre: p.nombre + ' (' + p.tamano + ')', imagen_url: p.imagenUrl, puntos_requeridos: p.puntosRequeridos, tipo: 'postre' }));

    return platos.concat(bebidas, postres).sort((a, b) => a.puntos_requeridos - b.puntos_requeridos);
  } catch (err) {
    return [];
  }
}

// 🆕 Canjea un premio - resta los puntos usados (no resetea el resto),
// y deja una fila en Pedidos para que aparezca en el panel de "Pedidos"
// como algo a preparar y entregar, igual que cualquier otro pedido.
function canjearPremio(idUsuario, idProducto) {
  try {
    const hojaUsuarios = obtenerOCrearHojaUsuarios();
    const filaUsuario = encontrarFilaUsuario(hojaUsuarios, idUsuario);
    if (!filaUsuario) return { ok: false, error: 'No se encontró tu usuario.' };

    const datosUsuario = hojaUsuarios.getRange(filaUsuario, 1, 1, ENCABEZADOS_USUARIOS.length).getValues()[0];
    const puntosActuales = Number(datosUsuario[10]) || 0;

    const premios = obtenerPremiosDisponibles();
    const premio = premios.find(p => String(p.id) === String(idProducto));
    if (!premio) return { ok: false, error: 'Ese premio ya no está disponible.' };
    if (puntosActuales < premio.puntos_requeridos) return { ok: false, error: 'No te alcanzan los puntos todavía.' };

    hojaUsuarios.getRange(filaUsuario, 11).setValue(puntosActuales - premio.puntos_requeridos);

    // Deja registrado el canje como un pedido más (total $0, aclarado en
    // el detalle) - así el panel de Pedidos avisa que hay que prepararlo
    // y entregarlo igual que cualquier otro.
    registrarPedido({
      idUsuario: idUsuario,
      nombreCliente: datosUsuario[1],
      telefono: datosUsuario[4],
      tipoDestino: 'retiro',
      empresaDestino: 'CANJE DE PREMIO - a coordinar entrega',
      fechaEntrega: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      horaEntrega: '--:--',
      detalleItems: '🎁 PREMIO CANJEADO: ' + premio.nombre,
      detalleItemsJson: JSON.stringify([{ tipo: 'premio', nombre: premio.nombre, variante: '', cantidad: 1 }]),
      formaPago: 'efectivo',
      total: 0
    });

    return { ok: true, puntosRestantes: puntosActuales - premio.puntos_requeridos };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// 🆕 Agrega o saca un producto de favoritos - alterna solo, sin
// necesidad de mandar el estado actual desde el frontend.
function toggleFavorito(idUsuario, idProducto) {
  try {
    const hoja = obtenerOCrearHojaUsuarios();
    const fila = encontrarFilaUsuario(hoja, idUsuario);
    if (!fila) return { ok: false, error: 'No se encontró tu usuario.' };

    const valorActual = hoja.getRange(fila, 13).getValue();
    let favoritos = valorActual ? String(valorActual).split(',').map(s => s.trim()).filter(Boolean) : [];

    if (favoritos.includes(String(idProducto))) {
      favoritos = favoritos.filter(id => id !== String(idProducto));
    } else {
      favoritos.push(String(idProducto));
    }

    hoja.getRange(fila, 13).setValue(favoritos.join(','));
    return { ok: true, favoritos: favoritos };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ==========================================================================
// 🆕 BEBIDAS Y POSTRES - antes vivían fijos en el código de menus.html,
// ahora se cargan desde el Sheets, con el mismo formato para las 2 hojas:
//   A: id
//   B: nombre       (ej: "Gaseosa", "Flan")
//   C: variantes    (separadas por coma - sabores, marcas, lo que sea:
//                    "Coca-Cola, Sprite, Fanta")
//   D: precio
//   E: tamano       (ej: "500ml", "180g")
//   F: disponible   ("si" / "no")
// Una fila por cada combinación de tamaño (el mismo "Gaseosa" en 500ml y
// en 2.25L son 2 filas separadas, cada una con su propio precio).
// ==========================================================================
const ENCABEZADOS_BEBIDAS_POSTRES = ["id", "nombre", "variantes", "precio", "tamano", "disponible", "es_premio", "puntos_requeridos", "imagen_url"];
const NOMBRE_HOJA_BEBIDAS = "Bebidas";
const NOMBRE_HOJA_POSTRES = "Postres";

function obtenerOCrearHojaBebidas() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_BEBIDAS);
  if (!hoja) {
    // 🆕 Reutiliza la hoja "Adicionales" si existe (para no dejarla sin
    // uso) - la renombra a "Bebidas" y le acomoda los encabezados.
    const hojaVieja = ss.getSheetByName("Adicionales");
    if (hojaVieja) {
      hojaVieja.setName(NOMBRE_HOJA_BEBIDAS);
      hoja = hojaVieja;
    } else {
      hoja = ss.insertSheet(NOMBRE_HOJA_BEBIDAS);
    }
  }
  const encabezadosActuales = hoja.getRange(1, 1, 1, ENCABEZADOS_BEBIDAS_POSTRES.length).getValues()[0];
  const desactualizada = ENCABEZADOS_BEBIDAS_POSTRES.some((titulo, i) => encabezadosActuales[i] !== titulo);
  if (desactualizada) {
    hoja.getRange(1, 1, 1, ENCABEZADOS_BEBIDAS_POSTRES.length).setValues([ENCABEZADOS_BEBIDAS_POSTRES]);
  }
  // Si la hoja quedó recién creada (o recién renombrada) sin ninguna
  // fila de datos todavía, la siembra con lo que ya venía funcionando,
  // para no perder disponibilidad de un día para el otro.
  if (hoja.getLastRow() < 2) {
    hoja.getRange(2, 1, 2, 9).setValues([
      ["1", "Agua", "Sin gas", 800, "500ml", "si", "no", "", ""],
      ["2", "Gaseosa", "Cola, Limón, Naranja, Manzana", 1000, "500ml", "si", "no", "", ""]
    ]);
    hoja.appendRow(["3", "Gaseosa", "Cola, Limón, Naranja, Manzana", 2200, "2.25L", "si", "no", "", ""]);
  }
  return hoja;
}

function obtenerOCrearHojaPostres() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_POSTRES);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_POSTRES);
  }
  const encabezadosActuales = hoja.getRange(1, 1, 1, ENCABEZADOS_BEBIDAS_POSTRES.length).getValues()[0];
  const desactualizada = ENCABEZADOS_BEBIDAS_POSTRES.some((titulo, i) => encabezadosActuales[i] !== titulo);
  if (desactualizada) {
    hoja.getRange(1, 1, 1, ENCABEZADOS_BEBIDAS_POSTRES.length).setValues([ENCABEZADOS_BEBIDAS_POSTRES]);
  }
  if (hoja.getLastRow() < 2) {
    hoja.getRange(2, 1, 3, 9).setValues([
      ["1", "Flan", "Dulce de leche, Chantilly, Mixto", 900, "180g", "si", "no", "", ""],
      ["2", "Postre", "Vainilla, Chocolate", 950, "180g", "si", "no", "", ""],
      ["3", "Ensalada de frutas", "Con chantilly, Sola", 800, "180g", "si", "no", "", ""]
    ]);
  }
  return hoja;
}

// Lee cualquiera de las 2 hojas (mismo formato) y arma la lista para el
// panel de admin (edición) o para la app del cliente (obtenerBebidas /
// obtenerPostres, más abajo, filtran solo lo disponible).
function leerHojaBebidasOPostres(hoja) {
  const filas = hoja.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < filas.length; i++) {
    const [id, nombre, variantes, precio, tamano, disponible, esPremio, puntosRequeridos, imagenUrl] = filas[i];
    if (!nombre) continue;
    items.push({
      fila: i + 1,
      id: String(id),
      nombre: String(nombre).trim(),
      variantes: variantes ? String(variantes).trim() : '',
      precio: Number(precio) || 0,
      tamano: tamano ? String(tamano).trim() : '',
      disponible: String(disponible).trim().toLowerCase() === 'si',
      esPremio: String(esPremio || 'no').trim().toLowerCase() === 'si',
      puntosRequeridos: Number(puntosRequeridos) || 0,
      imagenUrl: imagenUrl ? String(imagenUrl).trim() : ''
    });
  }
  return items;
}

function obtenerBebidasAdmin() { return leerHojaBebidasOPostres(obtenerOCrearHojaBebidas()); }
function obtenerPostresAdmin() { return leerHojaBebidasOPostres(obtenerOCrearHojaPostres()); }

// 🆕 Para la app del cliente - solo lo disponible, formato liviano.
function obtenerBebidas() { return obtenerBebidasAdmin().filter(b => b.disponible); }
function obtenerPostresCliente() { return obtenerPostresAdmin().filter(p => p.disponible); }

function guardarItemBebidaOPostre(hoja, datos) {
  try {
    const fila = datos.fila ? Number(datos.fila) : (hoja.getLastRow() + 1);

    // 🆕 Si mandaron una foto nueva, la sube y usa esa URL - si no,
    // mantiene la que ya estaba guardada (para no perderla al editar
    // solo el precio, por ejemplo).
    let imagenUrl = datos.imagenUrlExistente || '';
    if (datos.imagenArchivo) {
      const idArchivo = subirImagenDrive(datos.imagenArchivo);
      if (idArchivo) imagenUrl = "https://drive.google.com/file/d/" + idArchivo + "/view";
    }

    hoja.getRange(fila, 1, 1, 9).setValues([[
      datos.id, datos.nombre, datos.variantes || '', Number(datos.precio) || 0,
      datos.tamano || '', datos.disponible || 'si',
      datos.esPremio || 'no', datos.puntosRequeridos || '', imagenUrl
    ]]);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function guardarBebida(datos) { return guardarItemBebidaOPostre(obtenerOCrearHojaBebidas(), datos); }
function guardarPostre(datos) { return guardarItemBebidaOPostre(obtenerOCrearHojaPostres(), datos); }

// ==========================================================================
// 🆕 CANJE RETROACTIVO - si el cliente se olvidó de sumar el premio al
// hacer un pedido, pero ese pedido es de HOY y todavía no fue
// entregado, se le puede agregar el premio ahí mismo (el viaje todavía
// no se hizo, así que no hace falta uno nuevo).
// ==========================================================================

// Trae los pedidos de HOY de un usuario que todavía no están
// entregados - candidatos para canje retroactivo.
function obtenerPedidosPendientesHoyDeUsuario(idUsuario) {
  try {
    const hoyTexto = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    return obtenerPedidosDeUsuario(idUsuario)
      .filter(p => p.fechaEntrega === hoyTexto && p.estadoPedido !== 'Entregado')
      .sort((a, b) => b.horaEntrega.localeCompare(a.horaEntrega));
  } catch (err) {
    return [];
  }
}

// Agrega uno o más premios a un pedido YA EXISTENTE (en vez de crear uno
// nuevo) - valida y descuenta los puntos igual que registrarPedido(),
// y actualiza el detalle del pedido para que el premio quede reflejado.
function agregarPremioAPedidoExistente(filaPedido, idUsuario, premios) {
  try {
    if (!premios || premios.length === 0) return { ok: false, error: 'No se especificó ningún premio.' };

    const hojaUsuarios = obtenerOCrearHojaUsuarios();
    const filaUsuario = encontrarFilaUsuario(hojaUsuarios, idUsuario);
    if (!filaUsuario) return { ok: false, error: 'No se encontró tu usuario.' };

    const puntosActuales = Number(hojaUsuarios.getRange(filaUsuario, 11).getValue()) || 0;
    const puntosNecesarios = premios.reduce((acc, p) => acc + (Number(p.puntosRequeridos) || 0), 0);
    if (puntosActuales < puntosNecesarios) {
      return { ok: false, error: 'No te alcanzan los puntos para ese/esos premio(s).' };
    }

    const hojaPedidos = obtenerOCrearHojaPedidos();
    // Confirma que el pedido sea de HOY, de ESTE usuario, y no esté ya
    // entregado - blindaje del lado del servidor, no solo confiar en lo
    // que mande el frontend.
    const filaCompleta = hojaPedidos.getRange(filaPedido, 1, 1, 15).getValues()[0];
    const [ , , idUsuarioPedido, , , , , fechaEntregaPedido, , detalleItemsActual, detalleItemsJsonActual, , , estadoPedidoActual ] = filaCompleta;
    const hoyTexto = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (String(idUsuarioPedido) !== String(idUsuario)) return { ok: false, error: 'Ese pedido no es tuyo.' };
    if (fechaEntregaPedido !== hoyTexto) return { ok: false, error: 'Ese pedido ya no es de hoy - no se puede agregar el premio.' };
    if (estadoPedidoActual === 'Entregado') return { ok: false, error: 'Ese pedido ya fue entregado - no se puede agregar el premio.' };

    // Descuenta los puntos.
    hojaUsuarios.getRange(filaUsuario, 11).setValue(puntosActuales - puntosNecesarios);

    // Suma los premios al detalle del pedido (texto + estructurado),
    // sin tocar lo que ya tenía.
    const detalleNuevo = detalleItemsActual + premios.map(p => ', 🎁 BENEFICIO: ' + p.nombre).join('');
    let itemsJson = [];
    try { itemsJson = JSON.parse(detalleItemsJsonActual || '[]'); } catch (e) { itemsJson = []; }
    premios.forEach(p => itemsJson.push({ tipo: 'premio', nombre: p.nombre, variante: '', cantidad: 1, puntos: p.puntosRequeridos }));

    hojaPedidos.getRange(filaPedido, 10).setValue(detalleNuevo); // J: detalle_items
    hojaPedidos.getRange(filaPedido, 11).setValue(JSON.stringify(itemsJson)); // K: detalle_items_json

    return { ok: true, puntosRestantes: puntosActuales - puntosNecesarios };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ==========================================================================
// 🆕 CONFIGURACIÓN - valores que el vendedor puede cambiar sin tocar
// código, editando una celda en el Sheets. Hoja "Configuracion":
//   A: clave       B: valor
// ==========================================================================
const NOMBRE_HOJA_CONFIGURACION = "Configuracion";

function obtenerOCrearHojaConfiguracion() {
  const ss = obtenerHojaCalculo();
  let hoja = ss.getSheetByName(NOMBRE_HOJA_CONFIGURACION);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_CONFIGURACION);
    hoja.appendRow(["clave", "valor"]);
    hoja.appendRow(["minutos_limite_sin_cargo", 30]);
    hoja.appendRow(["recargo_cancelacion_tardia", 500]);
  }
  return hoja;
}

function obtenerValorConfig(clave, valorDefault) {
  try {
    const hoja = obtenerOCrearHojaConfiguracion();
    const filas = hoja.getDataRange().getValues();
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]).trim() === clave) return Number(filas[i][1]) || valorDefault;
    }
    return valorDefault;
  } catch (err) {
    return valorDefault;
  }
}

// ==========================================================================
// 🆕 CANCELAR PEDIDO - distinto de borrarlo de la hoja. Si pasó más del
// límite configurado desde que se hizo el pedido, le queda un recargo
// pendiente al cliente, que se le suma solo al PRÓXIMO pedido que haga
// (no se le puede cobrar el actual, ya está cancelado).
// ==========================================================================
function cancelarPedido(fila) {
  try {
    const hojaPedidos = obtenerOCrearHojaPedidos();
    const filaCompleta = hojaPedidos.getRange(fila, 1, 1, 15).getValues()[0];
    const [ , fechaHora, idUsuario, , , , , , , , detalleItemsJson, , , estadoPedidoActual ] = filaCompleta;

    if (estadoPedidoActual === 'Entregado') {
      return { ok: false, error: 'Ese pedido ya fue entregado - no se puede cancelar.' };
    }
    if (estadoPedidoActual === 'Cancelado') {
      return { ok: false, error: 'Ese pedido ya estaba cancelado.' };
    }

    // 🆕 Si el pedido incluía algún premio canjeado con puntos, esos
    // puntos se devuelven a la cuenta del cliente al cancelar - no
    // tiene sentido que se queden gastados si el pedido no se entrega.
    let puntosDevueltos = 0;
    try {
      const items = JSON.parse(detalleItemsJson || '[]');
      puntosDevueltos = items.filter(i => i.tipo === 'premio').reduce((acc, i) => acc + (Number(i.puntos) || 0), 0);
    } catch (e) { puntosDevueltos = 0; }

    if (puntosDevueltos > 0 && idUsuario && idUsuario !== 'INVITADO') {
      const bloqueoPuntos = LockService.getScriptLock();
      bloqueoPuntos.waitLock(10000);
      try {
        const hojaUsuarios = obtenerOCrearHojaUsuarios();
        const filaUsuario = encontrarFilaUsuario(hojaUsuarios, idUsuario);
        if (filaUsuario) {
          const puntosActuales = Number(hojaUsuarios.getRange(filaUsuario, 11).getValue()) || 0;
          hojaUsuarios.getRange(filaUsuario, 11).setValue(puntosActuales + puntosDevueltos);
        }
      } finally {
        bloqueoPuntos.releaseLock();
      }
    }

    const minutosLimite = obtenerValorConfig('minutos_limite_sin_cargo', 30);
    const montoRecargo = obtenerValorConfig('recargo_cancelacion_tardia', 500);
    const minutosTranscurridos = (new Date().getTime() - new Date(fechaHora).getTime()) / 60000;

    let llevaRecargo = false;
    if (minutosTranscurridos > minutosLimite && idUsuario && idUsuario !== 'INVITADO') {
      // 🔧 FIX RAÍZ: si se cancelan 2 pedidos casi al mismo tiempo, sin
      // este bloqueo las 2 llamadas pueden leer el MISMO valor de
      // recargo_pendiente antes de que la primera termine de guardar -
      // la segunda pisa el resultado en vez de sumarse (por eso solo
      // quedaba UN recargo cargado en vez de los 2). El bloqueo obliga a
      // que una espere a que la otra termine, para leer-sumar-guardar
      // sin que se crucen.
      const bloqueo = LockService.getScriptLock();
      bloqueo.waitLock(10000); // espera hasta 10 segundos si otra cancelación está en curso
      try {
        const hojaUsuarios = obtenerOCrearHojaUsuarios();
        const filaUsuario = encontrarFilaUsuario(hojaUsuarios, idUsuario);
        if (filaUsuario) {
          const recargoActual = Number(hojaUsuarios.getRange(filaUsuario, 14).getValue()) || 0;
          hojaUsuarios.getRange(filaUsuario, 14).setValue(recargoActual + montoRecargo);
          llevaRecargo = true;
        }
      } finally {
        bloqueo.releaseLock();
      }
    }

    hojaPedidos.getRange(fila, 14).setValue('Cancelado'); // N: estado_pedido

    return { ok: true, llevaRecargo: llevaRecargo, monto: montoRecargo, puntosDevueltos: puntosDevueltos };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
