
const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs');
/** vbs実行用（同期的にexe実行を行う）*/
const { spawnSync } = require('child_process')

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })

  win.loadFile('index.html')
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

/** VBS へ処理を渡す */
ipcMain.on('excel-test', function(event, args){
  console.log('main!')

  /** VBSファイルのディレクトリ（コンパイル後か否かで変える） */
  let dir = './resources/app/vbs'
  if (!fs.existsSync(dir)) dir = './vbs'

  /** 実行するVBSファイルのパス */
  var fullpath = `${dir}/excel.vbs`;

  /** VBS に処理を渡す */
  var child = spawnSync('cscript.exe', [fullpath, args]);

  /** 戻り値を取得する(status) */
  var ret = child.status;

  /** 戻り値の確認 */
  if (ret === 0) {
    dialog.showMessageBox({
      type: 'info',
      title: '通知',
      message: '実行確認',
      detail: '処理が実行されました',
      button:['OK'],
    })
  } else {
    console.log('error!')
  }
});
