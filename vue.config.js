const { join: joinPath } = require('path'); // eslint-disable-line @typescript-eslint/no-var-requires

function resolve(dir) {
  return joinPath(__dirname, dir);
}

module.exports = {
  publicPath: '/',
  configureWebpack: {
    entry: {
      app: './demo/main.ts',
    },
    resolve: {
      alias: {
        holysheetjs: resolve('./src'),
      },
    },
  },
  chainWebpack: config => {
    config.plugin('html').tap(args => {
      args[0].template = resolve('./demo/index.html');

      return args;
    });
  },
  devServer: {
    disableHostCheck: true,
  },
};
