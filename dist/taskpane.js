/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={27091:function(t){"use strict";t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},44944:function(t,e,n){"use strict";t.exports=n.p+"assets/logo-filled.png"},36076:function(t,e,n){"use strict";t.exports=n.p+"f8a72f06229e96e9b49d.js"},60806:function(t,e,n){"use strict";t.exports=n.p+"1fda685b81e1123773f6.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var a=e[r]={exports:{}};return t[r](a,a.exports,n),a.exports}n.m=t,n.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return n.d(e,{a:e}),e},n.d=function(t,e){for(var r in e)n.o(e,r)&&!n.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");r.length&&(t=r[r.length-1].src)}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(t,e){var n="undefined"!=typeof Symbol&&t[Symbol.iterator]||t["@@iterator"];if(!n){if(Array.isArray(t)||(n=o(t))||e&&t&&"number"==typeof t.length){n&&(t=n);var r=0,a=function(){};return{s:a,n:function(){return r>=t.length?{done:!0}:{done:!1,value:t[r++]}},e:function(t){throw t},f:a}}throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}var c,u=!0,i=!1;return{s:function(){n=n.call(t)},n:function(){var t=n.next();return u=t.done,t},e:function(t){i=!0,c=t},f:function(){try{u||null==n.return||n.return()}finally{if(i)throw c}}}}function n(){"use strict";n=function(){return e};var e={},r=Object.prototype,o=r.hasOwnProperty,a=Object.defineProperty||function(t,e,n){t[e]=n.value},c="function"==typeof Symbol?Symbol:{},u=c.iterator||"@@iterator",i=c.asyncIterator||"@@asyncIterator",s=c.toStringTag||"@@toStringTag";function f(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(t){f=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var o=e&&e.prototype instanceof v?e:v,c=Object.create(o.prototype),u=new _(r||[]);return a(c,"_invoke",{value:E(t,n,u)}),c}function p(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=l;var h={};function v(){}function y(){}function d(){}var m={};f(m,u,(function(){return this}));var g=Object.getPrototypeOf,w=g&&g(g(I([])));w&&w!==r&&o.call(w,u)&&(m=w);var b=d.prototype=v.prototype=Object.create(m);function x(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function k(e,n){function r(a,c,u,i){var s=p(e[a],e,c);if("throw"!==s.type){var f=s.arg,l=f.value;return l&&"object"==t(l)&&o.call(l,"__await")?n.resolve(l.__await).then((function(t){r("next",t,u,i)}),(function(t){r("throw",t,u,i)})):n.resolve(l).then((function(t){f.value=t,u(f)}),(function(t){return r("throw",t,u,i)}))}i(s.arg)}var c;a(this,"_invoke",{value:function(t,e){function o(){return new n((function(n,o){r(t,e,n,o)}))}return c=c?c.then(o,o):o()}})}function E(t,e,n){var r="suspendedStart";return function(o,a){if("executing"===r)throw new Error("Generator is already running");if("completed"===r){if("throw"===o)throw a;return{value:void 0,done:!0}}for(n.method=o,n.arg=a;;){var c=n.delegate;if(c){var u=L(c,n);if(u){if(u===h)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if("suspendedStart"===r)throw r="completed",n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);r="executing";var i=p(t,e,n);if("normal"===i.type){if(r=n.done?"completed":"suspendedYield",i.arg===h)continue;return{value:i.arg,done:n.done}}"throw"===i.type&&(r="completed",n.method="throw",n.arg=i.arg)}}}function L(t,e){var n=e.method,r=t.iterator[n];if(void 0===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=void 0,L(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),h;var o=p(r,t.iterator,e.arg);if("throw"===o.type)return e.method="throw",e.arg=o.arg,e.delegate=null,h;var a=o.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=void 0),e.delegate=null,h):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,h)}function S(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function O(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function _(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(S,this),this.reset(!0)}function I(t){if(t){var e=t[u];if(e)return e.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var n=-1,r=function e(){for(;++n<t.length;)if(o.call(t,n))return e.value=t[n],e.done=!1,e;return e.value=void 0,e.done=!0,e};return r.next=r}}return{next:j}}function j(){return{value:void 0,done:!0}}return y.prototype=d,a(b,"constructor",{value:d,configurable:!0}),a(d,"constructor",{value:y,configurable:!0}),y.displayName=f(d,s,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===y||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,d):(t.__proto__=d,f(t,s,"GeneratorFunction")),t.prototype=Object.create(b),t},e.awrap=function(t){return{__await:t}},x(k.prototype),f(k.prototype,i,(function(){return this})),e.AsyncIterator=k,e.async=function(t,n,r,o,a){void 0===a&&(a=Promise);var c=new k(l(t,n,r,o),a);return e.isGeneratorFunction(n)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},x(b),f(b,s,"Generator"),f(b,u,(function(){return this})),f(b,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=I,_.prototype={constructor:_,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=void 0,this.done=!1,this.delegate=null,this.method="next",this.arg=void 0,this.tryEntries.forEach(O),!t)for(var e in this)"t"===e.charAt(0)&&o.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=void 0)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,r){return c.type="throw",c.arg=t,e.next=n,r&&(e.method="next",e.arg=void 0),!!r}for(var r=this.tryEntries.length-1;r>=0;--r){var a=this.tryEntries[r],c=a.completion;if("root"===a.tryLoc)return n("end");if(a.tryLoc<=this.prev){var u=o.call(a,"catchLoc"),i=o.call(a,"finallyLoc");if(u&&i){if(this.prev<a.catchLoc)return n(a.catchLoc,!0);if(this.prev<a.finallyLoc)return n(a.finallyLoc)}else if(u){if(this.prev<a.catchLoc)return n(a.catchLoc,!0)}else{if(!i)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return n(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&o.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var a=r;break}}a&&("break"===t||"continue"===t)&&a.tryLoc<=e&&e<=a.finallyLoc&&(a=null);var c=a?a.completion:{};return c.type=t,c.arg=e,a?(this.method="next",this.next=a.finallyLoc,h):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),h},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),O(n),h}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;O(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:I(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=void 0),h}},e}function r(t){return function(t){if(Array.isArray(t))return a(t)}(t)||function(t){if("undefined"!=typeof Symbol&&null!=t[Symbol.iterator]||null!=t["@@iterator"])return Array.from(t)}(t)||o(t)||function(){throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}()}function o(t,e){if(t){if("string"==typeof t)return a(t,e);var n=Object.prototype.toString.call(t).slice(8,-1);return"Object"===n&&t.constructor&&(n=t.constructor.name),"Map"===n||"Set"===n?Array.from(t):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?a(t,e):void 0}}function a(t,e){(null==e||e>t.length)&&(e=t.length);for(var n=0,r=new Array(e);n<e;n++)r[n]=t[n];return r}function c(t,e,n,r,o,a,c){try{var u=t[a](c),i=u.value}catch(t){return void n(t)}u.done?e(i):Promise.resolve(i).then(r,o)}function u(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var a=t.apply(e,n);function u(t){c(a,r,o,u,i,"next",t)}function i(t){c(a,r,o,u,i,"throw",t)}u(void 0)}))}}var i="_BDD",s="baseEtapes",f="baseParents",l="Données_entrée",p="Configuration - Entrées Sorties",h="tableConfig";function v(){return y.apply(this,arguments)}function y(){return y=u(n().mark((function t(){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:Excel.run(function(){var t=u(n().mark((function t(e){var r,o,a,c;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,E();case 3:return t.next=5,g();case 5:return r=t.sent,t.next=8,e.sync();case 8:return o=r[0],a=r[1],t.next=12,d(o);case 12:return c=t.sent,o.forEach(function(){var t=u(n().mark((function t(e){var r,u;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:1==e[0]?_(e,c[0]):(r=e[0],u=a.filter((function(t){return t[1]==r})),j(e,c[r-1],u,o,c));case 1:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),t.next=16,e.sync();case 16:return t.abrupt("return",e);case 19:return t.prev=19,t.t0=t.catch(0),t.t0,Office.context.ui.displayDialogAsync(window.location.origin+"/error.html",{height:45,width:55},(function(t){var e=t.value;e.addEventHandler(Office.EventType.DialogMessageReceived,(function(t){JSON.parse(t.message),e.close()}))})),t.abrupt("return",e.sync());case 23:case"end":return t.stop()}}),t,null,[[0,19]])})));return function(e){return t.apply(this,arguments)}}());case 1:case"end":return t.stop()}}),t)}))),y.apply(this,arguments)}function d(t){return m.apply(this,arguments)}function m(){return m=u(n().mark((function t(r){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,o){Excel.run(function(){var o=u(n().mark((function o(a){var c,u,s,f,l,p,h,v,y,d;return n().wrap((function(n){for(;;)switch(n.prev=n.next){case 0:c=[],u=a.workbook.worksheets.getItem(i),s=e(r),n.prev=4,s.s();case 6:if((f=s.n()).done){n.next=28;break}return l=f.value,p=l[0],h=l[1],v=h+"|"+p,n.next=13,x(v);case 13:if(n.sent){n.next=23;break}return y=a.workbook.worksheets.getItem("MODEL_"+h),n.next=18,a.sync();case 18:return d=y.copy("Before",u),n.next=21,a.sync();case 21:d.name=v,d.visibility="Visible";case 23:return n.next=25,a.sync();case 25:c.push(v);case 26:n.next=6;break;case 28:n.next=33;break;case 30:n.prev=30,n.t0=n.catch(4),s.e(n.t0);case 33:return n.prev=33,s.f(),n.finish(33);case 36:return n.next=38,a.sync();case 38:S(c),t(c);case 40:case"end":return n.stop()}}),o,null,[[4,30,33,36]])})));return function(t){return o.apply(this,arguments)}}()).catch((function(t){return o(t)}))})));case 1:case"end":return t.stop()}}),t)}))),m.apply(this,arguments)}function g(){return w.apply(this,arguments)}function w(){return w=u(n().mark((function t(){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,e){Excel.run(function(){var e=u(n().mark((function e(r){var o,a,c;return n().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return o=r.workbook.worksheets.getItem(i),a=o.tables.getItem(s).getRange().getUsedRange(),c=o.tables.getItem(f).getRange().getUsedRange(),a.load("values"),c.load("values"),e.next=7,r.sync();case 7:a.values.shift(),c.values.shift(),b(a.values,c.values),t([a.values,c.values]);case 11:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()).catch((function(t){return e(t)}))})));case 1:case"end":return t.stop()}}),t)}))),w.apply(this,arguments)}function b(t,e){if(0==t.length)throw new Error("La base d'étapes est vide");if(0==e.length)throw new Error("La base de parents est vide");var n=t.map((function(t){return t[0]})),o=r(new Set(n));if(n.length!=o.length)throw console.log(n,o),new Error("Les id des étapes ne sont pas uniques");var a=e.map((function(t){return t[0]})),c=e.map((function(t){return t[1]})),u=a.map((function(t,e){return[t,c[e]]})),i=r(new Set(u));if(u.length!=i.length)throw new Error("Les id des parents et enfants ne sont pas uniques");r(new Set(a)).forEach((function(t){if(100!=e.filter((function(e){return e[0]==t})).reduce((function(t,e){return t+e[2]}),0))throw new Error("La somme des flux pour l'étape ".concat(t," est différente de 100"))}))}function x(t){return k.apply(this,arguments)}function k(){return k=u(n().mark((function t(e){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,r){Excel.run(function(){var r=u(n().mark((function r(o){var a,c;return n().wrap((function(n){for(;;)switch(n.prev=n.next){case 0:return a=o.workbook.worksheets,c=a.getItemOrNullObject(e),n.next=4,o.sync();case 4:t(!c.isNullObject);case 5:case"end":return n.stop()}}),r)})));return function(t){return r.apply(this,arguments)}}()).catch((function(t){return r(t)}))})));case 1:case"end":return t.stop()}}),t)}))),k.apply(this,arguments)}function E(){return L.apply(this,arguments)}function L(){return L=u(n().mark((function t(){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Excel.run(function(){var t=u(n().mark((function t(e){var r,o,a,c,u;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=e.workbook.worksheets.getItem(i),o=r.tables.getItem(s),a=r.tables.getItem(f),c=[{key:0,ascending:!0}],u=[{key:0,ascending:!0}],o.sort.apply(c),a.sort.apply(u),t.next=9,e.sync();case 9:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),L.apply(this,arguments)}function S(t){return O.apply(this,arguments)}function O(){return O=u(n().mark((function t(e){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Excel.run(function(){var t=u(n().mark((function t(r){var o,a,c;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(o=r.workbook.worksheets).load("items/name"),t.next=4,r.sync();case 4:for(a=0;a<o.items.length;a++)(c=o.items[a]).name.includes("|")&&!e.includes(c.name)&&c.delete();return t.next=7,r.sync();case 7:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),O.apply(this,arguments)}function _(t,e){return I.apply(this,arguments)}function I(){return I=u(n().mark((function t(e,r){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=3,Excel.run(function(){var t=u(n().mark((function t(o){var a,c,u,i,s,f,v;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return a=e[1],c=o.workbook.worksheets.getItem(r),u=o.workbook.worksheets.getItem(l),c.load("values"),u.load("values"),t.next=7,o.sync();case 7:return t.next=9,D(p,h,"DONNEES_ENTREES");case 9:return i=t.sent,t.next=12,D(p,h,a+"_Entrée");case 12:if(s=t.sent,i[0].length===s[0].length){t.next=16;break}throw console.log(i[0]+" vs "+s[0]),new Error("Les colonnes données d'entrées et colonnesEtapeUneEntree n'ont pas la même longueur : \n");case 16:for(f=1;f<i[0].length;f++)for(v=0;v<4;v++)c.getRange(s[v][f][0]).values=[["=".concat(l,"!").concat(i[v][f][0])]];return t.next=19,o.sync();case 19:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:case"end":return t.stop()}}),t)}))),I.apply(this,arguments)}function j(t,e,n,r){return P.apply(this,arguments)}function P(){return P=u(n().mark((function t(e,r,o,a){return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=3,Excel.run(function(){var t=u(n().mark((function t(c){var i,s,f,l,v,y,d,m,g;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return i=e[1],e[0],s=c.workbook.worksheets.getItem(r),t.next=5,c.sync();case 5:return s.load("values"),t.next=8,c.sync();case 8:return f=[],o.forEach(function(){var t=u(n().mark((function t(e){var r,o,c;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=a.find((function(t){return t[0]===e[0]}))[1],o=r+"|"+e[0],t.next=4,G(o,r+"_Sortie");case 4:c=t.sent,f.push(c);case 6:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),t.next=12,G(r,i+"_Entree");case 12:return l=t.sent,t.next=15,D(p,h,"TYPE_DE_CHAMP");case 15:if(v=t.sent,l.length==f[0].length&&l.length==v[0].length){t.next=19;break}throw console.log(l[0]+" vs "+f[0][0]),new Error("Les colonnes target et colonnes de tabSources n'ont pas la même longueur : \n");case 19:y=1;case 20:if(!(y<l.length)){t.next=42;break}d=0;case 22:if(!(d<4)){t.next=39;break}m=l[y][d].split("!"),g=s.getRange(m[1]),t.t0=v[0][y][0],t.next="Débit"===t.t0?28:"Concentration"===t.t0?30:"Température"===t.t0?32:"PH"===t.t0?34:36;break;case 28:return g.formulas=[[R(f,y,d,o)]],t.abrupt("break",36);case 30:return g.formulas=[[N(f,y,d,o)]],t.abrupt("break",36);case 32:return g.formulas=[[T(f,y,d,o)]],t.abrupt("break",36);case 34:return g.formulas=[[A(f,y,d)]],t.abrupt("break",36);case 36:d++,t.next=22;break;case 39:y++,t.next=20;break;case 42:return t.next=44,c.sync();case 44:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:case"end":return t.stop()}}),t)}))),P.apply(this,arguments)}function R(t,e,n,r){for(var o="=",a=0;a<r.length;a++)o+="(".concat(t[a][e][n],"*").concat(r[a][2],"/100)+");return o.slice(0,-1)}function N(t,e,n,r){for(var o="=(",a=0;a<r.length;a++)o+="(".concat(t[a][e][0],"*").concat(t[a][e][n],"*").concat(r[a][2],"/100)+");o=o.slice(0,-1),o+=") / (";for(var c=0;c<r.length;c++)o+="(".concat(t[c][e][0],"*").concat(r[c][2],"/100)+");return(o=o.slice(0,-1))+")"}function T(t,e,n,r){for(var o="=",a=0;a<r.length;a++)o+="(".concat(t[a][e][n],"*").concat(r[a][2],"/100)+");return o.slice(0,-1)}function A(t,e,n,r){return"=".concat(t[0][e][n])}function D(t,e,n){return U.apply(this,arguments)}function U(){return U=u(n().mark((function t(e,o,a){var c;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return c=[],t.next=3,Excel.run(function(){var t=u(n().mark((function t(u){var i,s,f,l,p,h;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return i=u.workbook.worksheets.getItem(e),(s=i.tables.getItem(o).getRange().getUsedRange()).load("values"),t.next=5,u.sync();case 5:return f=s.values[0],l=f.map((function(t,e){return t.startsWith(a)?e:-1})).filter((function(t){return t>=0})),t.next=9,Promise.all(l.map((function(t){return s.getColumn(t).getUsedRange().load("values")})));case 9:return p=t.sent,t.next=12,u.sync();case 12:h=p.map((function(t){return t.values})),c.push.apply(c,r(h));case 14:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:return t.abrupt("return",c);case 4:case"end":return t.stop()}}),t)}))),U.apply(this,arguments)}function G(t,e){return B.apply(this,arguments)}function B(){return B=u(n().mark((function t(e,r){var o;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,o=[],t.next=4,Excel.run(function(){var t=u(n().mark((function t(a){var c,u,i,s,f,l,p,h,v,y;return n().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(c=a.workbook.worksheets.getItem(e)).load("tables"),t.next=4,a.sync();case 4:return(u=c.tables).load("items/name"),t.next=8,a.sync();case 8:if(i=u.items.find((function(t){return t.name.startsWith(r)}))){t.next=12;break}return console.error('Table with prefix "'.concat(r,'" not found')),t.abrupt("return",null);case 12:return(s=i.getDataBodyRange()).load("address"),s.load("values/length"),t.next=17,a.sync();case 17:f=s.values.length,l=s.values[0].length,p=0;case 20:if(!(p<f)){t.next=36;break}h=[],v=0;case 23:if(!(v<l)){t.next=32;break}return(y=s.getCell(p,v)).load("address"),t.next=28,a.sync();case 28:h.push(y.address);case 29:v++,t.next=23;break;case 32:o.push(h);case 33:p++,t.next=20;break;case 36:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 4:return t.abrupt("return",o);case 7:return t.prev=7,t.t0=t.catch(0),console.error(t.t0),t.abrupt("return",null);case 11:case"end":return t.stop()}}),t,null,[[0,7]])}))),B.apply(this,arguments)}Office.onReady((function(t){t.host===Office.HostType.Excel&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("app").addEventListener("click",v))}))}(),function(){"use strict";var t=n(27091),e=n.n(t),r=new URL(n(60806),n.b),o=new URL(n(36076),n.b),a=new URL(n(44944),n.b);e()(r),e()(o),e()(a)}()}();
//# sourceMappingURL=taskpane.js.map