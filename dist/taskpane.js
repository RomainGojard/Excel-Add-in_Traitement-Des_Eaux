/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var t,e,n,r,o={27091:function(t){t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},44944:function(t,e,n){t.exports=n.p+"assets/logo-filled.png"},60806:function(t,e,n){t.exports=n.p+"1fda685b81e1123773f6.css"}},a={};function c(t){var e=a[t];if(void 0!==e)return e.exports;var n=a[t]={exports:{}};return o[t](n,n.exports,c),n.exports}c.m=o,c.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return c.d(e,{a:e}),e},c.d=function(t,e){for(var n in e)c.o(e,n)&&!c.o(t,n)&&Object.defineProperty(t,n,{enumerable:!0,get:e[n]})},c.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),c.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;c.g.importScripts&&(t=c.g.location+"");var e=c.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");n.length&&(t=n[n.length-1].src)}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),c.p=t}(),c.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(t,e){var r="undefined"!=typeof Symbol&&t[Symbol.iterator]||t["@@iterator"];if(!r){if(Array.isArray(t)||(r=n(t))||e&&t&&"number"==typeof t.length){r&&(t=r);var o=0,a=function(){};return{s:a,n:function(){return o>=t.length?{done:!0}:{done:!1,value:t[o++]}},e:function(t){throw t},f:a}}throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}var c,u=!0,i=!1;return{s:function(){r=r.call(t)},n:function(){var t=r.next();return u=t.done,t},e:function(t){i=!0,c=t},f:function(){try{u||null==r.return||r.return()}finally{if(i)throw c}}}}function n(t,e){if(t){if("string"==typeof t)return r(t,e);var n=Object.prototype.toString.call(t).slice(8,-1);return"Object"===n&&t.constructor&&(n=t.constructor.name),"Map"===n||"Set"===n?Array.from(t):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?r(t,e):void 0}}function r(t,e){(null==e||e>t.length)&&(e=t.length);for(var n=0,r=new Array(e);n<e;n++)r[n]=t[n];return r}function o(){o=function(){return e};var e={},n=Object.prototype,r=n.hasOwnProperty,a=Object.defineProperty||function(t,e,n){t[e]=n.value},c="function"==typeof Symbol?Symbol:{},u=c.iterator||"@@iterator",i=c.asyncIterator||"@@asyncIterator",s=c.toStringTag||"@@toStringTag";function f(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(t){f=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var o=e&&e.prototype instanceof y?e:y,c=Object.create(o.prototype),u=new _(r||[]);return a(c,"_invoke",{value:E(t,n,u)}),c}function p(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=l;var h={};function y(){}function v(){}function d(){}var m={};f(m,u,(function(){return this}));var g=Object.getPrototypeOf,w=g&&g(g(O([])));w&&w!==n&&r.call(w,u)&&(m=w);var b=d.prototype=y.prototype=Object.create(m);function x(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function k(e,n){function o(a,c,u,i){var s=p(e[a],e,c);if("throw"!==s.type){var f=s.arg,l=f.value;return l&&"object"==t(l)&&r.call(l,"__await")?n.resolve(l.__await).then((function(t){o("next",t,u,i)}),(function(t){o("throw",t,u,i)})):n.resolve(l).then((function(t){f.value=t,u(f)}),(function(t){return o("throw",t,u,i)}))}i(s.arg)}var c;a(this,"_invoke",{value:function(t,e){function r(){return new n((function(n,r){o(t,e,n,r)}))}return c=c?c.then(r,r):r()}})}function E(t,e,n){var r="suspendedStart";return function(o,a){if("executing"===r)throw new Error("Generator is already running");if("completed"===r){if("throw"===o)throw a;return{value:void 0,done:!0}}for(n.method=o,n.arg=a;;){var c=n.delegate;if(c){var u=L(c,n);if(u){if(u===h)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if("suspendedStart"===r)throw r="completed",n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);r="executing";var i=p(t,e,n);if("normal"===i.type){if(r=n.done?"completed":"suspendedYield",i.arg===h)continue;return{value:i.arg,done:n.done}}"throw"===i.type&&(r="completed",n.method="throw",n.arg=i.arg)}}}function L(t,e){var n=e.method,r=t.iterator[n];if(void 0===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=void 0,L(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),h;var o=p(r,t.iterator,e.arg);if("throw"===o.type)return e.method="throw",e.arg=o.arg,e.delegate=null,h;var a=o.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=void 0),e.delegate=null,h):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,h)}function I(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function S(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function _(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(I,this),this.reset(!0)}function O(t){if(t){var e=t[u];if(e)return e.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var n=-1,o=function e(){for(;++n<t.length;)if(r.call(t,n))return e.value=t[n],e.done=!1,e;return e.value=void 0,e.done=!0,e};return o.next=o}}return{next:j}}function j(){return{value:void 0,done:!0}}return v.prototype=d,a(b,"constructor",{value:d,configurable:!0}),a(d,"constructor",{value:v,configurable:!0}),v.displayName=f(d,s,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===v||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,d):(t.__proto__=d,f(t,s,"GeneratorFunction")),t.prototype=Object.create(b),t},e.awrap=function(t){return{__await:t}},x(k.prototype),f(k.prototype,i,(function(){return this})),e.AsyncIterator=k,e.async=function(t,n,r,o,a){void 0===a&&(a=Promise);var c=new k(l(t,n,r,o),a);return e.isGeneratorFunction(n)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},x(b),f(b,s,"Generator"),f(b,u,(function(){return this})),f(b,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=O,_.prototype={constructor:_,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=void 0,this.done=!1,this.delegate=null,this.method="next",this.arg=void 0,this.tryEntries.forEach(S),!t)for(var e in this)"t"===e.charAt(0)&&r.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=void 0)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,r){return c.type="throw",c.arg=t,e.next=n,r&&(e.method="next",e.arg=void 0),!!r}for(var o=this.tryEntries.length-1;o>=0;--o){var a=this.tryEntries[o],c=a.completion;if("root"===a.tryLoc)return n("end");if(a.tryLoc<=this.prev){var u=r.call(a,"catchLoc"),i=r.call(a,"finallyLoc");if(u&&i){if(this.prev<a.catchLoc)return n(a.catchLoc,!0);if(this.prev<a.finallyLoc)return n(a.finallyLoc)}else if(u){if(this.prev<a.catchLoc)return n(a.catchLoc,!0)}else{if(!i)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return n(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var a=o;break}}a&&("break"===t||"continue"===t)&&a.tryLoc<=e&&e<=a.finallyLoc&&(a=null);var c=a?a.completion:{};return c.type=t,c.arg=e,a?(this.method="next",this.next=a.finallyLoc,h):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),h},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),S(n),h}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;S(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:O(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=void 0),h}},e}function a(t,e,n,r,o,a,c){try{var u=t[a](c),i=u.value}catch(t){return void n(t)}u.done?e(i):Promise.resolve(i).then(r,o)}function c(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var c=t.apply(e,n);function u(t){a(c,r,o,u,i,"next",t)}function i(t){a(c,r,o,u,i,"throw",t)}u(void 0)}))}}var u="_BDD",i="baseEtapes",s="baseParents",f="Données_entrée",l="Configuration - Entrées Sorties",p="tableConfig";function h(){return y.apply(this,arguments)}function y(){return y=c(o().mark((function t(){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:Excel.run(function(){var t=c(o().mark((function t(e){var n,r,a,u;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,x();case 3:return t.next=5,m();case 5:return n=t.sent,t.next=8,e.sync();case 8:return r=n[0],a=n[1],t.next=12,v(r);case 12:return u=t.sent,r.forEach(function(){var t=c(o().mark((function t(e){var n,c;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:1==e[0]?I(e,u[0]):(n=e[0],c=a.filter((function(t){return t[1]==n})),_(e,u[n-1],c,r,u));case 1:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),t.next=16,e.sync();case 16:return t.abrupt("return",e);case 19:return t.prev=19,t.t0=t.catch(0),console.error(t.t0),t.abrupt("return",e.sync());case 23:case"end":return t.stop()}}),t,null,[[0,19]])})));return function(e){return t.apply(this,arguments)}}());case 1:case"end":return t.stop()}}),t)}))),y.apply(this,arguments)}function v(t){return d.apply(this,arguments)}function d(){return d=c(o().mark((function t(n){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,r){Excel.run(function(){var r=c(o().mark((function r(a){var c,i,s,f,l,p,h,y,v,d;return o().wrap((function(r){for(;;)switch(r.prev=r.next){case 0:c=[],i=a.workbook.worksheets.getItem(u),s=e(n),r.prev=4,s.s();case 6:if((f=s.n()).done){r.next=28;break}return l=f.value,p=l[0],h=l[1],y=h+"|"+p,r.next=13,w(y);case 13:if(r.sent){r.next=23;break}return v=a.workbook.worksheets.getItem("MODEL_"+h),r.next=18,a.sync();case 18:return d=v.copy("Before",i),r.next=21,a.sync();case 21:d.name=y,d.visibility="Visible";case 23:return r.next=25,a.sync();case 25:c.push(y);case 26:r.next=6;break;case 28:r.next=33;break;case 30:r.prev=30,r.t0=r.catch(4),s.e(r.t0);case 33:return r.prev=33,s.f(),r.finish(33);case 36:return r.next=38,a.sync();case 38:E(c),t(c);case 40:case"end":return r.stop()}}),r,null,[[4,30,33,36]])})));return function(t){return r.apply(this,arguments)}}()).catch((function(t){return r(t)}))})));case 1:case"end":return t.stop()}}),t)}))),d.apply(this,arguments)}function m(){return g.apply(this,arguments)}function g(){return g=c(o().mark((function t(){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,e){Excel.run(function(){var e=c(o().mark((function e(n){var r,a,c;return o().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem(u),a=r.tables.getItem(i).getRange().getUsedRange(),c=r.tables.getItem(s).getRange().getUsedRange(),a.load("values"),c.load("values"),e.next=7,n.sync();case 7:a.values.shift(),c.values.shift(),t([a.values,c.values]);case 10:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()).catch((function(t){return e(t)}))})));case 1:case"end":return t.stop()}}),t)}))),g.apply(this,arguments)}function w(t){return b.apply(this,arguments)}function b(){return b=c(o().mark((function t(e){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,n){Excel.run(function(){var n=c(o().mark((function n(r){var a,c;return o().wrap((function(n){for(;;)switch(n.prev=n.next){case 0:return a=r.workbook.worksheets,c=a.getItemOrNullObject(e),n.next=4,r.sync();case 4:t(!c.isNullObject);case 5:case"end":return n.stop()}}),n)})));return function(t){return n.apply(this,arguments)}}()).catch((function(t){return n(t)}))})));case 1:case"end":return t.stop()}}),t)}))),b.apply(this,arguments)}function x(){return k.apply(this,arguments)}function k(){return k=c(o().mark((function t(){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Excel.run(function(){var t=c(o().mark((function t(e){var n,r,a,c,f;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=e.workbook.worksheets.getItem(u),r=n.tables.getItem(i),a=n.tables.getItem(s),c=[{key:0,ascending:!0}],f=[{key:0,ascending:!0}],r.sort.apply(c),a.sort.apply(f),t.next=9,e.sync();case 9:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),k.apply(this,arguments)}function E(t){return L.apply(this,arguments)}function L(){return L=c(o().mark((function t(e){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Excel.run(function(){var t=c(o().mark((function t(n){var r,a,c;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(r=n.workbook.worksheets).load("items/name"),t.next=4,n.sync();case 4:for(a=0;a<r.items.length;a++)(c=r.items[a]).name.includes("|")&&!e.includes(c.name)&&c.delete();return t.next=7,n.sync();case 7:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),L.apply(this,arguments)}function I(t,e){return S.apply(this,arguments)}function S(){return S=c(o().mark((function t(e,n){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=3,Excel.run(function(){var t=c(o().mark((function t(r){var a,c,u,i,s,h,y;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return a=e[1],c=r.workbook.worksheets.getItem(n),u=r.workbook.worksheets.getItem(f),c.load("values"),u.load("values"),t.next=7,r.sync();case 7:return t.next=9,N(l,p,"DONNEES_ENTREES");case 9:return i=t.sent,t.next=12,N(l,p,a+"_Entrée");case 12:if(s=t.sent,i[0].length===s[0].length){t.next=16;break}throw console.log(i[0]+" vs "+s[0]),new Error("Les colonnes données d'entrées et colonnesEtapeUneEntree n'ont pas la même longueur : \n");case 16:for(h=1;h<i[0].length;h++)for(y=0;y<4;y++)c.getRange(s[y][h][0]).values=[["=".concat(f,"!").concat(i[y][h][0])]];return t.next=19,r.sync();case 19:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:case"end":return t.stop()}}),t)}))),S.apply(this,arguments)}function _(t,e,n,r){return O.apply(this,arguments)}function O(){return O=c(o().mark((function t(e,n,r,a){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=3,Excel.run(function(){var t=c(o().mark((function t(u){var i,s,f,h,y,v,d,m;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return i=e[1],e[0],s=u.workbook.worksheets.getItem(n),t.next=5,u.sync();case 5:return s.load("values"),t.next=8,u.sync();case 8:return t.next=10,D(n,i+"_Entree");case 10:return f=[],r.forEach(function(){var t=c(o().mark((function t(e){var n,r,c;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=a.find((function(t){return t[0]===e[0]}))[1],r=n+"|"+e[0],t.next=4,N(l,p,n+"_Sortie");case 4:(c=t.sent).push(r),f.push(c);case 7:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()),t.next=14,N(l,p,i+"_Entrée");case 14:return h=t.sent,t.next=17,N(l,p,"TYPE_DE_CHAMP");case 17:if(y=t.sent,h[0].length===f[0][0].length||h[0].length===y[0].length||y[0].length===f[0][0].length){t.next=21;break}throw console.log(h[0]+" vs "+f[0][0]),new Error("Les colonnes target et colonnes de tabSources n'ont pas la même longueur : \n");case 21:console.log(f),console.log(h),console.log(y[0]),v=1;case 25:if(!(v<h[0].length)){t.next=46;break}d=0;case 27:if(!(d<4)){t.next=43;break}m=s.getRange(h[d][v][0]),t.t0=y[0][v][0],t.next="Débit"===t.t0?32:"Concentration"===t.t0?34:"Température"===t.t0?36:"PH"===t.t0?38:40;break;case 32:return m.formulas=[[j(f,d,v,r)]],t.abrupt("break",40);case 34:return m.formulas=[[P(f,d,v,r)]],t.abrupt("break",40);case 36:return m.formulas=[[T(f,d,v,r)]],t.abrupt("break",40);case 38:return m.formulas=[[R(f,d,v)]],t.abrupt("break",40);case 40:d++,t.next=27;break;case 43:v++,t.next=25;break;case 46:return t.next=48,u.sync();case 48:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:case"end":return t.stop()}}),t)}))),O.apply(this,arguments)}function j(t,e,n,r){for(var o="=",a=0;a<r.length;a++)o+="('".concat(t[a].slice(-1),"'!").concat(t[a][e][n][0],"*").concat(r[a][2],"/100)+");return o.slice(0,-1)}function P(t,e,n,r){for(var o="=(",a=0;a<r.length;a++)o+="('".concat(t[a].slice(-1),"'!").concat(t[a][e][1][0],"*'").concat(t[a].slice(-1),"'!").concat(t[a][e][n][0],"*").concat(r[a][2],"/100)+");o=o.slice(0,-1),o+=") / (";for(var c=0;c<r.length;c++)o+="('".concat(t[c].slice(-1),"'!").concat(t[c][e][1][0],"*").concat(r[c][2],"/100)+");return(o=o.slice(0,-1))+")"}function T(t,e,n,r){for(var o="=",a=0;a<r.length;a++)o+="('".concat(t[a].slice(-1),"'!").concat(t[a][e][n][0],"*").concat(r[a][2],"/100)+");return o.slice(0,-1)}function R(t,e,n,r){return"='".concat(t[0].slice(-1),"'!").concat(t[0][e][n][0])}function N(t,e,n){return A.apply(this,arguments)}function A(){return A=c(o().mark((function t(e,a,u){var i;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return i=[],t.next=3,Excel.run(function(){var t=c(o().mark((function t(c){var s,f,l,p,h,y;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return s=c.workbook.worksheets.getItem(e),(f=s.tables.getItem(a).getRange().getUsedRange()).load("values"),t.next=5,c.sync();case 5:return l=f.values[0],p=l.map((function(t,e){return t.startsWith(u)?e:-1})).filter((function(t){return t>=0})),t.next=9,Promise.all(p.map((function(t){return f.getColumn(t).getUsedRange().load("values")})));case 9:return h=t.sent,t.next=12,c.sync();case 12:y=h.map((function(t){return t.values})),i.push.apply(i,function(t){if(Array.isArray(t))return r(t)}(o=y)||function(t){if("undefined"!=typeof Symbol&&null!=t[Symbol.iterator]||null!=t["@@iterator"])return Array.from(t)}(o)||n(o)||function(){throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}());case 14:case"end":return t.stop()}var o}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:return t.abrupt("return",i);case 4:case"end":return t.stop()}}),t)}))),A.apply(this,arguments)}function D(t,e){return B.apply(this,arguments)}function B(){return B=c(o().mark((function t(e,n){return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,Excel.run(function(){var t=c(o().mark((function t(r){var a,c,u,i;return o().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(a=r.workbook.worksheets.getItem(e)).load("tables"),t.next=4,r.sync();case 4:return(c=a.tables).load("items/name"),t.next=8,r.sync();case 8:if(u=c.items.find((function(t){return t.name.startsWith(n)}))){t.next=12;break}return console.error('Table with prefix "'.concat(n,'" not found')),t.abrupt("return",null);case 12:return(i=u.getDataBodyRange()).load("values"),t.next=16,r.sync();case 16:console.log(i.values);case 17:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 3:t.next=8;break;case 5:t.prev=5,t.t0=t.catch(0),console.error(t.t0);case 8:case"end":return t.stop()}}),t,null,[[0,5]])}))),B.apply(this,arguments)}Office.onReady((function(t){t.host===Office.HostType.Excel&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("open-dialog").onclick=G,document.getElementById("app").addEventListener("click",h))}));var U=null;function G(){Office.context.ui.displayDialogAsync("https://csb10032001a6800cf9.z6.web.core.windows.net/popup.html",{height:45,width:55},(function(t){(U=t.value).addEventHandler(Office.EventType.DialogMessageReceived,M)}))}function M(t){document.getElementById("user-name").innerHTML=t.message,U.close()}}(),t=c(27091),e=c.n(t),n=new URL(c(60806),c.b),r=new URL(c(44944),c.b),e()(n),e()(r)}();
//# sourceMappingURL=taskpane.js.map