
const {Scheduler, XTTReader} = require('../lib/index')

async function testReader () {
  let xtt = new XTTReader()
  await xtt.loadXTT('test/DotNetSample_signed.xtt')
  console.log(xtt.getNumOfTests())
  console.log(xtt.getTestInfo(0))
  // let keys = xtt.generateRSAKeys()
  // await xtt.signXTTFile('test/DotNetSample.xtt', 'test/DotNetSample_signed.xtt', keys[1])
}

async function testScheduler () {
  let object = new Scheduler()
  object.on('FOLDER_END', (msg) => {
    console.log('TEST_RESULT', msg)
  })

  let tree = await object.openXTT('test/DotNetSample.xtt')
  let r = await tree.runTree()
  console.log('ff:', r)

  console.log(tree.getLoopRowCount('Looping'))
  await tree.stopTest()
  console.log('stopped')
  await object.exit()
  console.log('closed')
  object.dispose()

//  tree = await object.openXTT('test/DotNetSample.xtt')
  // r = await tree.runTree()
  // console.log('ff:', r)
}

// testReader()

testScheduler()
testScheduler()
