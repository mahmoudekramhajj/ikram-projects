/**
 * State.js — إدارة Script Properties
 *
 * استخدام:
 *   GDS2.State.get('key');
 *   GDS2.State.set('key', value);
 *   GDS2.State.getJSON('IATA_REGISTRY');
 *   GDS2.State.setJSON('IATA_REGISTRY', list);
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.State = {
  _props: function() {
    return PropertiesService.getScriptProperties();
  },

  get: function(key) {
    return GDS2.State._props().getProperty(key);
  },

  set: function(key, value) {
    GDS2.State._props().setProperty(key, String(value));
  },

  getJSON: function(key) {
    var v = GDS2.State.get(key);
    if (!v) return null;
    try {
      return JSON.parse(v);
    } catch (e) {
      GDS2.Log.error('State.getJSON parse error', { key: key, error: e.message });
      return null;
    }
  },

  setJSON: function(key, obj) {
    GDS2.State.set(key, JSON.stringify(obj));
  },

  remove: function(key) {
    GDS2.State._props().deleteProperty(key);
  },

  all: function() {
    return GDS2.State._props().getProperties();
  },

  // فحص: هل المفاتيح المطلوبة موجودة؟
  checkRequired: function(keys) {
    var missing = [];
    for (var i = 0; i < keys.length; i++) {
      if (!GDS2.State.get(keys[i])) {
        missing.push(keys[i]);
      }
    }
    return missing;
  }
};
