const { createProxyMiddleware } = require('http-proxy-middleware');

module.exports = function(app) {
  app.use(
    '/excel-processor',
    createProxyMiddleware({
      target: 'http://localhost:3000',
      pathRewrite: {
        '^/excel-processor': ''
      }
    })
  );
}; 