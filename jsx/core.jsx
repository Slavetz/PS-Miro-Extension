var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _jsonStringify = function () {
  var cx = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
      escapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
      gap,
      indent,
      meta = {
    '\b': '\\b',
    '\t': '\\t',
    '\n': '\\n',
    '\f': '\\f',
    '\r': '\\r',
    '"': '\\"',
    '\\': '\\\\'
  },
      rep;

  function quote(string) {
    escapable.lastIndex = 0;
    return escapable.test(string) ? '"' + string.replace(escapable, function (a) {
      var c = meta[a];
      return typeof c === 'string' ? c : '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
    }) + '"' : '"' + string + '"';
  }

  function str(key, holder) {
    var i,
        k,
        v,
        length,
        mind = gap,
        partial,
        value = holder[key];

    if (value && (typeof value === 'undefined' ? 'undefined' : _typeof(value)) === 'object' && typeof value.toJSON === 'function') {
      value = value.toJSON(key);
    }

    if (typeof rep === 'function') {
      value = rep.call(holder, key, value);
    }

    switch (typeof value === 'undefined' ? 'undefined' : _typeof(value)) {
      case 'string':
        return quote(value);

      case 'number':
        return isFinite(value) ? String(value) : 'null';

      case 'boolean':
      case 'null':
        return String(value);

      case 'object':
        if (!value) return 'null';
        gap += indent;
        partial = [];

        if (Object.prototype.toString.apply(value) === '[object Array]') {
          length = value.length;

          for (i = 0; i < length; i += 1) {
            partial[i] = str(i, value) || 'null';
          }

          v = partial.length === 0 ? '[]' : (gap ? '[\n' + gap + partial.join(',\n' + gap) + '\n' + mind + ']' : '[' + partial.join(',') + ']');
          gap = mind;
          return v;
        }

        if (rep && (typeof rep === 'undefined' ? 'undefined' : _typeof(rep)) === 'object') {
          length = rep.length;

          for (i = 0; i < length; i += 1) {
            k = rep[i];

            if (typeof k === 'string') {
              v = str(k, value);

              if (v) {
                partial.push(quote(k) + (gap ? ': ' : ':') + v);
              }
            }
          }
        } else {
          for (k in value) {
            if (Object.prototype.hasOwnProperty.call(value, k)) {
              v = str(k, value);

              if (v) {
                partial.push(quote(k) + (gap ? ': ' : ':') + v);
              }
            }
          }
        }

        v = partial.length === 0 ? '{}' : (gap ? '{\n' + gap + partial.join(',\n' + gap) + '\n' + mind + '}' : '{' + partial.join(',') + '}');
        gap = mind;
        return v;
    }
  }

  return function (value, replacer, space) {
    var i;
    gap = '';
    indent = '';

    if (typeof space === 'number') {
      for (i = 0; i < space; i += 1) {
        indent += ' ';
      }
    } else if (typeof space === 'string') {
      indent = space;
    }

    rep = replacer;

    if (replacer && typeof replacer !== 'function' && ((typeof replacer === 'undefined' ? 'undefined' : _typeof(replacer)) !== 'object' || typeof replacer.length !== 'number')) {
      throw new Error('JSON.stringify');
    }

    return str('', {
      '': value
    });
  };
}();

var _map = function _map(array, callback) {
  var newArray = [];

  for (var i = 0; i < array.length; i++) {
    newArray.push(callback(array[i], i, array));
  }

  return newArray;
};

var _filter = function _filter(array, callback) {
  var newArray = [];

  for (var i = 0; i < array.length; i++) {
    if (callback(array[i], i, array)) newArray.push(array[i]);
  }

  return newArray;
};

var _forEach = function _forEach(array, callback) {
  for (var i = 0; i < array.length; i++) {
    callback(array[i], i, array);
  }
};

function testFunction() {
  var savePath = File.saveDialog('Save path', '*.miro');
  if (savePath === null) return;
  return decodeURI(savePath.relativeURI);
}

function getPages(options) {
  var scale = options.scale;

  /** Получаем страницы */

  var pages = getSLidesObjectFromInDesign(scale);
  if (pages === undefined) return;

  /** Диалог сохранения файла */
  var savePath = File.saveDialog('Save path', '*.miro');
  if (savePath === null) return;

  return _jsonStringify({ path: decodeURI(savePath.relativeURI), pages: pages });
}

function exportPages(options) {
  var pages = options.pages,
      path = options.path,
      scale = options.scale;

  /** Пишем все картинки */

  exportImagesFromInDesign(pages, path, scale);

  exportJSON(pages, path);

  return true;
}

/**+++++++++++++++++++++++++++++++++++++++++++++++++**/
/**+++++++++++++++++++++++++++++++++++++++++++++++++**/

function getSLidesObjectFromInDesign(scale) {

  var document = app.activeDocument;
  var layername = 'prop';

  var propLayer = document.layers.itemByName(layername);

  if (!propLayer.isValid) {
    throw new Error('!!!Prop Layer is absent!!!');
  }

  var hmu = document.viewPreferences.horizontalMeasurementUnits;
  var vmu = document.viewPreferences.verticalMeasurementUnits;

  document.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PIXELS;
  document.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PIXELS;

  var pages = document.pages;
  var propLayerItems = document.layers.itemByName(layername).pageItems;

  var propSlides = _filter(propLayerItems, function (el) {
    return el.name === 'slide_prop_table';
  });

  var mod = 1;
  var cur_sld = '';
  var prevSpredWidth = 0;
  var currSpredWidth = 0;

  var slides_obj = _map(propSlides, function (item) {

    var content = item.tables[0].columns[1].contents;
    var e = item.parent.index;
    var el = pages[e];
    var sld = content[0].replace(/\s+/, '');

    if (cur_sld !== sld) {
      mod = 1;
      prevSpredWidth = prevSpredWidth + currSpredWidth + 500 * scale;
      currSpredWidth = 0;
    } else {
      mod += 1;
    }
    cur_sld = sld;

    var h = (el.bounds[2] - el.bounds[0]) * scale;
    var w = (el.bounds[3] - el.bounds[1]) * scale;
    if (currSpredWidth < el.bounds[3] * scale) {
      currSpredWidth = el.bounds[3] * scale;
    }

    return {
      title: 'Slide ' + sld + ' : Page ' + mod,
      externalId: el.id,
      filename: el.id + '.jpg',
      pageStr: el.appliedAlternateLayout.alternateLayout + ':' + (e + 1),
      y: h * 0.5 + h * 1.2 * mod,
      x: w * 0.5 + el.bounds[1] * scale + prevSpredWidth,
      height: h,
      width: w
    };
  });

  slides_obj = _filter(slides_obj, function (el) {
    return el !== undefined;
  });

  document.viewPreferences.horizontalMeasurementUnits = hmu;
  document.viewPreferences.verticalMeasurementUnits = vmu;

  return slides_obj;
}

function exportImagesFromInDesign(pages, path, scale) {

  _forEach(pages, function (page) {

    var file = new File(path + page.filename);

    app.jpegExportPreferences.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
    app.jpegExportPreferences.pageString = page.pageStr;

    app.jpegExportPreferences.exportResolution = 72 * scale;
    app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.HIGH;
    //	app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM;

    document.exportFile(ExportFormat.JPG, file);
  });
}

/**+++++++++++++++++++++++++++++++++++++++++++++++++**/
/**+++++++++++++++++++++++++++++++++++++++++++++++++**/

function exportJSON(json, saveStr) {

  var file = new File(saveStr + 'images.json');

  file.encoding = "utf-8";
  file.open("w");

  var jsonFileStr = _jsonStringify(json);

  file.write(jsonFileStr);
  file.close();

  return null;
}

function coreAlert(message) {
  alert(message);
}
