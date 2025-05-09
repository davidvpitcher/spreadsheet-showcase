const { defineConfig } = require('@vue/cli-service')

const isElectron = process.env.IS_ELECTRON === 'true'

module.exports = defineConfig({
  transpileDependencies: true,
  publicPath: isElectron ? './' : '/spreadsheet_test/',
  pluginOptions: {
    electronBuilder: {
      nodeIntegration: true,
      // Optional: tweak other Electron build settings here
    }
  }
})
