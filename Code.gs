/**
 * Registro Institucional de Bomberos — Tercera Compañía de Bomberos de Carelmapu
 * Backend Google Apps Script + Google Sheets + Drive.
 *
 * Instalación:
 * 1) Crear un proyecto Apps Script vinculado a una hoja de cálculo.
 * 2) Pegar este archivo como Code.gs y crear Index.html con el archivo entregado.
 * 3) Ejecutar setup() manualmente y autorizar permisos.
 * 4) Implementar > Nueva implementación > Aplicación web.
 *    Ejecutar como: tú. Acceso: según política de la institución.
 *
 * Seguridad: las contraseñas se guardan con hash SHA-256 + sal; sesiones temporales
 * en CacheService. Para una institución real, restringir acceso a usuarios autorizados
 * y revisar las políticas de Google Workspace antes de publicar.
 */
const APP_NAME = 'Registro de Bomberos Carelmapu';
const SHEETS = {
  Bomberos: ['id','nombres','apellidos','rut','fechaNacimiento','fechaIngreso','tipoBombero','situacion','codigoRadial','cargoInstitucional','telefono','email','direccion','grupoSanguineo','contactoEmergencia','telefonoEmergencia','observaciones','fotoUrl','createdAt','updatedAt'],
  Academico: ['id','bomberoId','fecha','tipo','cursoEspecialidad','institucion','resultado','horas','observaciones'],
  Material: ['id','bomberoId','fechaEntrega','categoria','material','marcaModelo','numeroSerie','estado','observaciones','fechaDevolucion'],
  Usuarios: ['id','usuario','nombre','passwordHash','salt','rol','activo','createdAt'],
  Auditoria: ['fecha','usuario','accion','entidad','registroId','detalle']
};
const ROLES = ['Administrador','Oficial','Consulta'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Incluye fragmentos HTML (CSS.html y JavaScript.html) dentro de Index.html. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Vincula este proyecto a una hoja de cálculo antes de ejecutar setup().');
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,SHEETS[name].length).setFontWeight('bold').setBackground('#dff5e8');
      sh.autoResizeColumns(1, SHEETS[name].length);
    }
  });
  let sh = ss.getSheetByName('Usuarios');
  if (sh.getLastRow() < 2) {
    const salt = Utilities.getUuid();
    sh.appendRow([Utilities.getUuid(),'admin','Administrador','',salt,'Administrador',true,new Date().toISOString()]);
    sh.getRange(sh.getLastRow(),4).setValue(hashPassword_('Cambiar123!', salt));
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('UPLOAD_FOLDER_ID')) {
    const folder = DriveApp.createFolder('Fotos - Registro Bomberos Carelmapu');
    props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  }
  return 'Configuración creada. Usuario inicial: admin / Cambiar123! — cambie la contraseña al ingresar.';
}

function api(payload) {
  payload = payload || {};
  try {
    const action = String(payload.action || '');
    if (action === 'login') return login_(payload);
    if (action === 'logout') return logout_(payload.token);
    if (action === 'bootstrap') {
      const user = requireUser_(payload.token);
      return ok_({user:user, bomberos: readRows_('Bomberos'), academico: readRows_('Academico'), material: readRows_('Material'), usuarios: user.rol === 'Administrador' ? readRows_('Usuarios').map(publicUser_) : []});
    }
    const user = requireUser_(payload.token);
    if (action === 'saveBombero') {
      requireRole_(user, ['Administrador','Oficial']);
      const row = normalizeBombero_(payload.data || {});
      const saved = upsert_('Bomberos', row);
      audit_(user, 'guardar', 'Bombero', saved.id, saved.nombres + ' ' + saved.apellidos);
      return ok_(saved);
    }
    if (action === 'deleteBombero') {
      requireRole_(user, ['Administrador']);
      deleteById_('Bomberos', payload.id);
      audit_(user, 'eliminar', 'Bombero', payload.id, '');
      return ok_({id:payload.id});
    }
    if (action === 'saveAcademico') {
      requireRole_(user, ['Administrador','Oficial']);
      const d = payload.data || {};
      if (!d.bomberoId || !d.tipo) throw new Error('Seleccione un bombero e indique el tipo de formación.');
      const row = pick_(SHEETS.Academico, Object.assign({id:d.id || Utilities.getUuid(), fecha:d.fecha || new Date().toISOString().slice(0,10)}, d));
      const saved = upsert_('Academico', row);
      audit_(user, 'guardar', 'Académico', saved.id, d.cursoEspecialidad || d.tipo);
      return ok_(saved);
    }
    if (action === 'deleteAcademico' || action === 'deleteMaterial') {
      requireRole_(user, ['Administrador','Oficial']);
      const sheet = action === 'deleteAcademico' ? 'Academico' : 'Material';
      deleteById_(sheet, payload.id);
      audit_(user, 'eliminar', sheet, payload.id, '');
      return ok_({id:payload.id});
    }
    if (action === 'saveMaterial') {
      requireRole_(user, ['Administrador','Oficial']);
      const d = payload.data || {};
      if (!d.bomberoId || !d.material) throw new Error('Seleccione un bombero e indique el material.');
      const row = pick_(SHEETS.Material, Object.assign({id:d.id || Utilities.getUuid(), fechaEntrega:d.fechaEntrega || new Date().toISOString().slice(0,10)}, d));
      const saved = upsert_('Material', row);
      audit_(user, 'guardar', 'Material', saved.id, d.material);
      return ok_(saved);
    }
    if (action === 'saveUser') {
      requireRole_(user, ['Administrador']);
      const d = payload.data || {};
      if (!d.usuario || !d.nombre || !d.rol) throw new Error('Complete usuario, nombre y rol.');
      if (!ROLES.includes(d.rol)) throw new Error('Rol no válido.');
      const existing = readRows_('Usuarios').find(x => String(x.usuario).toLowerCase() === String(d.usuario).toLowerCase() && x.id !== d.id);
      if (existing) throw new Error('Ese nombre de usuario ya existe.');
      let row;
      if (d.id) {
        const old = readRows_('Usuarios').find(x => x.id === d.id);
        if (!old) throw new Error('Usuario no encontrado.');
        row = Object.assign({}, old, {nombre:d.nombre, rol:d.rol, activo:d.activo !== false});
        if (d.password) { row.salt = Utilities.getUuid(); row.passwordHash = hashPassword_(d.password, row.salt); }
      } else {
        if (!d.password || d.password.length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres.');
        row = {id:Utilities.getUuid(), usuario:d.usuario, nombre:d.nombre, salt:Utilities.getUuid(), rol:d.rol, activo:true, createdAt:new Date().toISOString()};
        row.passwordHash = hashPassword_(d.password, row.salt);
      }
      const saved = upsert_('Usuarios', row);
      audit_(user, 'guardar', 'Usuario', saved.id, saved.usuario);
      return ok_(publicUser_(saved));
    }
    if (action === 'uploadPhoto') {
      requireRole_(user, ['Administrador','Oficial']);
      const d = payload.data || {};
      if (!d.base64 || !d.fileName) throw new Error('No se recibió una fotografía válida.');
      const bytes = Utilities.base64Decode(d.base64);
      if (bytes.length > 5 * 1024 * 1024) throw new Error('La fotografía supera el límite de 5 MB.');
      const mime = String(d.mimeType || 'image/jpeg');
      if (!/^image\/(jpeg|png|webp)$/.test(mime)) throw new Error('Formato permitido: JPG, PNG o WEBP.');
      const folderId = PropertiesService.getScriptProperties().getProperty('UPLOAD_FOLDER_ID');
      const folder = DriveApp.getFolderById(folderId);
      const file = folder.createFile(Utilities.newBlob(bytes, mime, d.fileName));
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      return ok_({url:file.getUrl(), id:file.getId()});
    }
    if (action === 'changePassword') {
      const d = payload.data || {};
      const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      const users = readRows_('Usuarios');
      const found = users.find(x => x.id === user.id);
      if (!found || hashPassword_(String(d.currentPassword || ''), found.salt) !== found.passwordHash) throw new Error('La contraseña actual no es correcta.');
      if (!d.newPassword || d.newPassword.length < 8) throw new Error('La nueva contraseña debe tener al menos 8 caracteres.');
      found.salt = Utilities.getUuid(); found.passwordHash = hashPassword_(d.newPassword, found.salt);
      upsert_('Usuarios', found);
      return ok_({changed:true});
    }
    throw new Error('Acción no reconocida.');
  } catch (e) {
    return {ok:false, error:e && e.message ? e.message : String(e)};
  }
}

function login_(p) {
  const username = String(p.usuario || '').trim();
  const password = String(p.password || '');
  const u = readRows_('Usuarios').find(x => String(x.usuario).toLowerCase() === username.toLowerCase());
  if (!u || String(u.activo) === 'false' || u.activo === false || hashPassword_(password, u.salt) !== u.passwordHash) throw new Error('Usuario o contraseña incorrectos.');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({id:u.id, usuario:u.usuario, nombre:u.nombre, rol:u.rol}), 21600);
  audit_({usuario:u.usuario}, 'inicio de sesión', 'Sesión', u.id, '');
  return ok_({token:token, user:publicUser_(u)});
}

function ok_(data) {
  return {
    ok: true,
    data: data
  };
}
function logout_(token) { if (token) CacheService.getScriptCache().remove('session_' + token); return ok_({loggedOut:true}); }
function requireUser_(token) {
  if (!token) throw new Error('Sesión expirada. Inicie sesión nuevamente.');
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) throw new Error('Sesión expirada. Inicie sesión nuevamente.');
  return JSON.parse(raw);
}
function requireRole_(user, roles) { if (!roles.includes(user.rol)) throw new Error('No tiene permisos para realizar esta acción.'); }
function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function readRows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.filter(r => r.some(v => v !== '' && v !== null)).map(r => {
    const o = {};
    headers.forEach((h,i) => { let v = r[i]; if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); o[h] = v; });
    return o;
  });
}
function upsert_(name, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const headers = SHEETS[name];
  const row = pick_(headers, obj);
  const rows = readRows_(name);
  const index = rows.findIndex(r => String(r.id) === String(row.id));
  const values = headers.map(h => row[h] === undefined ? '' : row[h]);
  if (index >= 0) sh.getRange(index + 2, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);
  return row;
}
function deleteById_(name, id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const rows = readRows_(name);
  const index = rows.findIndex(r => String(r.id) === String(id));
  if (index >= 0) sh.deleteRow(index + 2);
}
function pick_(headers, obj) { const out = {}; headers.forEach(h => { if (obj[h] !== undefined) out[h] = obj[h]; }); if (!out.id) out.id = Utilities.getUuid(); return out; }
function normalizeBombero_(d) {
  if (!d.nombres || !d.apellidos) throw new Error('Nombres y apellidos son obligatorios.');
  const row = pick_(SHEETS.Bomberos, Object.assign({id:d.id || Utilities.getUuid(), createdAt:d.createdAt || new Date().toISOString(), updatedAt:new Date().toISOString()}, d));
  if (!row.situacion) row.situacion = 'Activo';
  if (!row.tipoBombero) row.tipoBombero = 'Bombero';
  return row;
}
function publicUser_(u) { return {id:u.id, usuario:u.usuario, nombre:u.nombre, rol:u.rol, activo:u.activo}; }
function audit_(user, action, entity, id, detail) {
  try { SpreadsheetApp.getActive().getSheetByName('Auditoria').appendRow([new Date().toISOString(), user.usuario || '', action, entity, id || '', detail || '']); } catch(e) {}
}
