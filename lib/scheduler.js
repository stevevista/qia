const ole = require('./bindings')
const TestTree = require('./testtree')
const { EventEmitter } = require('events')

const TestMsgEvents = [
  'UNIT_START',
  'UNIT_END',
  'TEST_START',
  'TEST_RESULT',
  'LOOP_START',
  'LOOP_END',
  'FOLDER_START',
  'FOLDER_END'
]

const ignoredAttrs = [
  // 'TestIndex',
  'TestFullPath',
  'RealName',
  'LoopDetails',
  'TestContainerName',
  'TestContainerVersion',
  'NodeRID'
]

class Scheduler extends EventEmitter {
  constructor (id) {
    super()
    this.id = id
    this.comp = new ole.Object('QSPR.Scheduler')
    this.tree = null
    this.dllpath = this.comp.__inprocServer32

    let that = this
    this.cookie = this.comp.callbackAdvise({
      __interface: 'ISchedulerEvents',
      OnDebugMessage: (...args) => that.onDebugMessage(...args),
      OnRealTimeParamMessage: (...args) => that.onRealTimeParamMessage(...args),
      OnTestMsg: (...args) => that.onTestMsg(...args),
      OnGlobalVariableChanged: (...args) => that.onGlobalVariableChanged(...args)
    })
  }

  dispose () {
    this.comp.callbackUnadvise(this.cookie)
  }

  onDebugMessage (strWin, strText, traceLevel, NoEndOfLine) {
    this.emit('DEBUG_MSG', strWin, strText, traceLevel, NoEndOfLine)
  }

  onRealTimeParamMessage (messageType, id, name, value, units, lowerLimit, upperLimit, pluginTypes, pluginNames) {
    // this event seems never fired
  }

  onTestMsg (messageType, messageData) {
    if (messageType > TestMsgEvents.length) {
      // should not happen
      return
    }

    const evtname = TestMsgEvents[messageType]
    let msg = { _id: this.id }
    const names = messageData.GetItemNames()
    for (let name of names) {
      if (ignoredAttrs.includes(name)) continue
      let val = messageData.GetValue(name)
      if (name === 'TestCount') val = parseInt(val)
      msg[name] = val
    }
    this.emit(evtname, msg)
  }

  onGlobalVariableChanged (gvName, gvValue) {
    this.emit('GLOBAL_VAR_CHANGE', gvName, gvValue)
  }

  openXTT (path) {
    // await this.close()
    return new Promise((resolve, reject) => {
      this.comp.OpenXTT(path, (err, ret) => {
        if (err) {
          reject(err)
        } else {
          this.tree = new TestTree(ret, path)
          resolve(this.tree)
        }
      })
    })
  }

  newXTT () {
    // await this.close()
    return new Promise((resolve, reject) => {
      this.comp.NewXTT((err, ret) => {
        if (err) {
          reject(err)
        } else {
          this.tree = new TestTree(ret, '')
          resolve(this.tree)
        }
      })
    })
  }

  openedXttPath () {
    return this.tree ? this.tree.path : null
  }

  lastError () {
    return this.comp.ErrorsExist ? this.comp.LastError : null
  }

  getGlobalVar (name) {
    let val = new ole.Variant('', 'pstring')
    this.comp.GetGlobalVariable(name, val)
    return val.valueOf()
  }

  setGlobalVar (name, val) {
    this.comp.SetGlobalVariable(name, val)
  }

  loadWorkspaceConfig (path) {
    return this._promiseWrap('LoadWorkspaceConfig', path)
  }

  saveWorkspaceConfig (path) {
    return this._promiseWrap('SaveWorkspaceConfig', path)
  }

  async close () {
    if (!this.tree) return
    await this._promiseWrap('Close')
    this.tree = null
  }

  exit () {
    return this._promiseWrap('Exit')
  }

  _promiseWrap (func, ...args) {
    return new Promise((resolve, reject) => {
      this.comp[func](...args, (err, ret) => {
        if (err) {
          reject(err)
        } else {
          resolve(ret)
        }
      })
    })
  }
}

module.exports = Scheduler
