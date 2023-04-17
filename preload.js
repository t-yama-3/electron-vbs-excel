const { ipcRenderer } = require('electron')

window.addEventListener('DOMContentLoaded', () => {
  const element = document.getElementById('excel-btn')
  element.addEventListener('click', () => {
    console.log('kokomade!')
    ipcRenderer.send('excel-test', 'dummy-args');
  })
})
