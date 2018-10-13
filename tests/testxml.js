const fs = require('fs')
//const { parseStructure, XTTStructParser } = require('../lib/index')
const { XTTStructParser } = require('../lib/structure')
function writeFile (path, data) {
  return new Promise((resolve, reject) => {
    fs.writeFile(path, data, (err) => {
      if (err) reject(err)
      else {
        resolve()
      }
    })
  })
}

function parseStructure1 (path) {
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
      console.log(props)
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

parseStructure1('C:\\Program Files (x86)\\Qualcomm\\QDART\\XTT\\SubSysGNSS\\GNSS_Gen8_LTE.xtt')

/*
parseStructure('test/DotNetSample.xtt')
  .then(ret => {
    console.log(ret)
    let xmlcontent = JSON.stringify(ret[0], null, 2)
    writeFile('./output11.yml', xmlcontent)
  })
// parser.parse('test/DotNetSample.xtt')
*/

