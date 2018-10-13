const ole = require('./bindings')

class XTTReader {
  constructor () {
    this.comp = new ole.Object('QSPR3XTT.reader')
  }

  loadXTT (path, publicKey) {
    return new Promise((resolve, reject) => {
      let func
      let args
      if (publicKey !== undefined) {
        func = 'LoadSignedQSPRTreeFile'
        args = [path, publicKey]
      } else {
        func = 'LoadQSPRTreeFile'
        args = [path]
      }
      this.comp[func](...args, (err, ret) => {
        if (err) reject(err)
        else {
          if (!ret) reject(new Error('LoadQSPRTreeFile return false'))
          else resolve()
        }
      })
    })
  }

  getNumOfTests () {
    return this.comp.GetNumOfTests()
  }

  getTestInfo (testIndex) {
    let testname = new ole.Variant('', 'pstring')
    let realname = new ole.Variant('', 'pstring')
    let numParams = new ole.Variant(0, 'pint')
    this.comp.GetTestInfo(testIndex, testname, realname, numParams)

    return [testname.valueOf(), realname.valueOf(), numParams.valueOf()]
  }

  getTestParameterInfo (testIndex, ParameterIndex) {
    let parameterName = new ole.Variant('', 'pstring')
    let parameterValue = new ole.Variant('', 'pstring')
    let unit = new ole.Variant('', 'pstring')
    let upperlimit = new ole.Variant('', 'pstring')
    let lowerlimit = new ole.Variant('', 'pstring')
    let DataType = new ole.Variant('', 'pstring')
    let paramMode = new ole.Variant('', 'pstring')
    this.comp.GetTestParameterInfo(testIndex, ParameterIndex, parameterName, parameterValue, unit, upperlimit, lowerlimit, DataType, paramMode)

    return {
      name: parameterName.valueOf(),
      value: parameterValue.valueOf(),
      unit: unit.valueOf(),
      upperlimit: upperlimit.valueOf(),
      lowerlimit: lowerlimit.valueOf(),
      type: DataType.valueOf(),
      mode: paramMode.valueOf()
    }
  }

  generateRSAKeys (seed) {
    if (seed === undefined) seed = 'node_qia'
    let publicKey = new ole.Variant('', 'pstring')
    let privateKey = new ole.Variant('', 'pstring')
    this.comp.GeneratePublicPrivateRSAKeys(seed, publicKey, privateKey)
    return [publicKey.valueOf(), privateKey.valueOf()]
  }

  signXTTFile (inFile, outFile, privateKey) {
    return this._promiseWrap('DigSignQSPRTreeFile', inFile, outFile, privateKey)
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

module.exports = XTTReader
