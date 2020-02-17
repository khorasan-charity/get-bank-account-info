'use strict'

const path = require('path')
const fs = require('fs')
const { VueLoaderPlugin } = require('vue-loader')

function resolve(dir) {
  return path.join(__dirname, '..', dir)
}

var jsEntries = {}
fs.readdirSync(resolve('./Views/')).forEach(function (name) {
    var indexFile = resolve('./Views/') + name + '/index.js'
    if (fs.existsSync(indexFile)) {
        jsEntries[name] = indexFile
    }
})

module.exports = {
  entry: jsEntries,
  output: {
      path: resolve('./wwwroot/js/dist'),
      publicPath: '/js/dist/',
      filename: '[name].js'
  },
  resolve: {
    extensions: ['.js', '.vue', '.json'],
    alias: {
      'vue$': 'vue/dist/vue.esm.js',
      '@': resolve('./Views/')
    }
  },

  module: {
    rules: [
      {
        test: /\.vue$/,
        use: 'vue-loader'
      }, {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
        }
      }
    ]
  },
  
  externals: {
    vue: 'Vue',
    quasar: 'Quasar',
    axios: 'axios'
  },

  plugins: [
    new VueLoaderPlugin()
  ]
}