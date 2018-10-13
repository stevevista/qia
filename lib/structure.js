const fs = require('fs')
const {EventEmitter} = require('events')
const sax = require('sax')

function stripBOM (str) {
  if (str[0] === '\uFEFF') {
    return str.substring(1)
  } else {
    return str
  }
}

class XTTStructParser extends EventEmitter {
  constructor () {
    super()
    this.options = {
      chunkSize: 100000
    }
    this.reset()
  }

  throwError (err) {
    if (!this.saxParser.errThrown) {
      this.saxParser.errThrown = true
      this.emit('error', err)
    }
  }

  reset () {
    this.removeAllListeners()
    this.saxParser = sax.parser(true, {
      trim: true,
      normalize: false
    })

    this.saxParser.errThrown = false
    this.saxParser.onerror = (error) => {
      this.saxParser.resume()
      this.throwError(error)
    }
    this.saxParser.onend = () => {
      this.emit('end')
    }

    let stack = []
    let props
    let text
    let looptype
    let variables =[]
    let pname

    this.saxParser.onopentag = (node) => {
      const tag = node.attributes['NodeType'] || node.name
      if (tag === 'Root' || tag === 'Folder' || tag === 'Test') {
        props = {}
        looptype = ''
        // if (tag === 'Test') props.Outputs = {}
      }
      text = ''
      stack.push(tag)
    }

    this.saxParser.onclosetag = () => {
      const tag = stack.pop()
      if (tag === 'Root' || tag === 'Folder') {
        this.emit('closefolder')
      } else if (tag === 'FolderProperty') {
        this.emit('folder', props)
      } else if (tag === 'TestProperty') {
        this.emit('test', props)
      } else if (tag === 'TestName') {
        const ptag = stack.length > 1 ? stack[stack.length - 2] : ''
        if (ptag === 'Root' || ptag === 'Folder' || ptag === 'Test') {
          if (ptag === 'Root') {
            if (props['TestName']) {
              this.emit('rootname', text)
            }
          }
          props['TestName'] = text
        }
      } else if (tag === 'Index') {
        props['TestIndex'] = text
      } else if (tag === 'CheckStateTest') {
        if (text === 'Unchecked') {
          props['Unchecked'] = true
        }
      } else if (tag === 'Expanded') {
        if (text === 'false') {
          props['Collapsed'] = true
        }
      } else if (tag === 'LoopType') {
        looptype = text
      } else if (tag === 'LoopValuesAsString') {
        let rounds = 1
        if (looptype === 'BYCOUNT') {
          rounds = parseInt(text)
        } else if (looptype === 'LIST') {
          rounds = text.split(',').length
        }
        props['LoopCount'] = rounds
      } else if (tag === 'Name') {
        const ptag = stack.length > 0 ? stack[stack.length - 1] : ''
        if (ptag === 'QGlobalVariable') {
          variables.push(text)
        } else if (ptag === 'ParameterCls') {
          // pname = text
        }
      } else if (tag === 'QGlobalVariables') {
        this.emit('variables', variables)
      } else if (tag === 'Type') {
        // if (text === 'Output') {
        //  props.Outputs[pname] = ''
        // }
      }
    }

    this.saxParser.ontext = (t) => {
      text += t
    }
  }

  processAsync () {
    try {
      if (this.remaining.length <= this.options.chunkSize) {
        const chunk = this.remaining
        this.remaining = ''
        this.saxParser.write(chunk).close()
      } else {
        const chunk = this.remaining.substr(0, this.options.chunkSize)
        this.remaining = this.remaining.substr(this.options.chunkSize, this.remaining.length)
        this.saxParser.write(chunk)
        setImmediate(() => this.processAsync())
      }
    } catch (err) {
      this.throwError(err)
    }
  }

  async parse (pathname) {
    let str = await readFile(pathname)
    try {
      str = str.toString()
      if (str.trim() === '') {
        this.emit('end')
        return
      }
      str = stripBOM(str)
      this.remaining = str
      setImmediate(() => this.processAsync())
    } catch (err) {
      this.throwError(err)
    }
  }
}

function readFile (path) {
  return new Promise((resolve, reject) => {
    fs.readFile(path, (err, data) => {
      if (err) reject(err)
      else {
        resolve(data)
      }
    })
  })
}

function parseStructure (path) {
  return new Promise((resolve, reject) => {
    const parser = new XTTStructParser()
    let outputs = []
    let stack = []
    let rootname

    parser.on('error', err => {
      reject(err)
    })

    parser.on('folder', props => {
      let parent
      if (stack.length > 0) {
        parent = stack[stack.length - 1]
      }
      stack.push(props)
      if (parent) {
        if (!parent.Tests) parent.Tests = [props]
        else parent.Tests.push(props)
      }
    })

    parser.on('closefolder', () => {
      let obj = stack.pop()
      if (stack.length === 0) {
        if (rootname) obj['TestName'] = rootname
        outputs.push(obj)
      }
    })

    parser.on('test', (props) => {
      let obj = stack[stack.length - 1]
      if (!obj.Tests) obj.Tests = [props]
      else obj.Tests.push(props)
    })

    parser.on('rootname', (name) => {
      rootname = name
    })

    parser.on('end', () => {
      resolve(outputs)
    })

    parser.parse(path)
  })
}

module.exports = {
  XTTStructParser,
  parseStructure
}
