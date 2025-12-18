const { Pdf4meExcel } = require('./dist/nodes/Pdf4me/Pdf4me.node.js');
const { Pdf4meExcelApi } = require('./dist/credentials/Pdf4meApi.credentials.js');

module.exports = {
  nodes: {
    Pdf4meExcel,
  },
  credentials: {
    Pdf4meExcelApi,
  },
};
