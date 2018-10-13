### Import QIA module

```
import {Scheduler, XTTReader} from 'qia'

```

## Scheduler
### Methods
* runTest
* stopTest

### Events
* **DEBUG_MSG**
```
scheduler.on('DEBUG_MSG', (win, text, level, NoEndOfLine) => ...)
```
* **UNIT_START**

测试开始(运行XTT测试)
```
scheduler.on('UNIT_START', msg => ...)
msg example:
{
  TestCount: 100,
  ...
}
```
* **TEST_RESULT**

单个测试项结束
```
scheduler.on('TEST_RESULT', msg => ...)
msg example:
{
  TestName: 'TestA',
  TestResult: 'Passed',
  OutputParam_MeasuredValue: '39.9002695641947',
  ...
}
```
