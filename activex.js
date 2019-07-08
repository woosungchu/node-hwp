var ActiveX = module.exports = require('./node-activex/node_activex');

global.ActiveXObject = function(id, opt) {
    return new ActiveX.Object(id, opt);
};
