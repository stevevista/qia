const ole = require('./bindings')

class TestTree {
  constructor (comp, path) {
    this.comp = comp
    this.path = path
  }

  save (pathname) {
    if (pathname) this.path = pathname
    return this._promiseWrap('SaveTree', pathname || '')
  }

  getTestNode (testname) {
    let ret = this.comp.GetTestNode(testname)
    if (ret) {
      return ret
    }
    return null
  }

  getNextTestInfo () {
    return this.comp.GetNextTestInfo()
  }

  lastError () {
    return this.comp.ErrorsExist ? this.comp.LastError : null
  }

  runTree (timeout) {
    return this._run(timeout, 'RunTree')
  }

  runTest (testName, timeout) {
    return this._run(timeout, 'RunTest', testName)
  }

  async _run (timeout, func, ...args) {
    let finished = false
    let result
    const cookie = this.comp.callbackAdvise({
      __interface: 'ITestTreeEvents',
      RunTreeDone (treePassed) {
        finished = true
        result = treePassed
      }
    })

    this.comp[func](...args)

    const stamp = Date.now()

    const waitFinish = () => {
      return new Promise((resolve, reject) => {
        const checkFinshFlag = () => {
          if (finished) {
            resolve()
          } else {
            if (timeout) {
              if (Date.now() - stamp > timeout) {
                reject(new Error('timeout'))
                return
              }
            }
            setImmediate(checkFinshFlag)
          }
        }
        setImmediate(checkFinshFlag)
      })
    }

    await waitFinish()
    this.comp.callbackUnadvise(cookie)

    return result
  }

  stopTest () {
    return this._promiseWrap('StopTest')
  }

  pauseTree () {
    return this._promiseWrap('PauseTree')
  }

  resumeTree () {
    return this._promiseWrap('ResumeTree')
  }

  runTreeInteractive () {
    return this._promiseWrap('RunTreeInteractive')
  }

  runNextTest () {
    return this._promiseWrap('RunNextTest')
  }

  selectTestToRun (testName, selectedToRun) {
    this.comp.SelectTestToRun(testName, selectedToRun)
  }

  selectTestFromFolderToRun (folder, testName, selectedToRun) {
    this.comp.SelectTestFromFolderToRun(folder, testName, selectedToRun)
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

  addDotNetDll (path) {
    return this.comp.AddDotNetDll(path)
  }

  getAllFolderTests (folderName) {
    let folders = new ole.Variant([], 'pstring')
    let folderCount = new ole.Variant(0, 'pint')
    this.comp.GetAllFolderTests(folderName, folders, folderCount)
    return folders.valueOf()
  }

  getAllChildTests (folderName) {
    let tests = new ole.Variant([], 'pstring')
    let folderCount = new ole.Variant(0, 'pint')
    this.comp.GetAllChildTests(folderName, tests, folderCount)
    return tests.valueOf()
  }

  getLoopRowCount (folderName) {
    // Q has a bug in this call
    try {
      return this.comp.GetLoopRowCount(folderName)
    } catch (e) {
      return 0
    }
  }
}

module.exports = TestTree
