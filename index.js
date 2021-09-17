import xlsx from 'node-xlsx'
import fs from "fs"
import { parse } from 'querystring'

const NOPES = ''

let EthiPattern = [
  { id: 'good', re: /^[\u4e00-\u9fa5]{1,2}族$/, f: x => x },
  { id: 'lack', re: /^[\u4e00-\u9fa5]{1,2}$/, f: x => x + '族' }
]
let YearMonth = [
  { id: 'good', re: /^[0-9]{4}年[0-9]{2}月$/, f: x => x },
  { id: 'extras', re: /^[0-9]{4}年[0-9]{2}月/, f: x => x.substr(0, 8) },
  { id: 'no月', re: /^[0-9]{4}年[0-9]{2}$/, f: x => x + '月' },
  { id: 'no月_short', re: /^[0-9]{4}年[0-9]{1}/, f: x => `${x.substr(0, 4)}年0${x[5]}月` },
  // { id: 'full_len', re: /[0-9]{8}/, f: x => `${x.substr(0, 4)}年${x.substr(4, 2)}月` },
  { id: 'half_len', re: /^[0-9]{6}/, f: x => `${x.substr(0, 4)}年${x.substr(4, 2)}月` },
  { id: 'hyphen', re: /^[0-9]{4}[－\- .～][0-9]{2}/, f: x => `${x.substr(0, 4)}年${x.substr(5, 2)}月` },
  { id: 'hyphen_short', re: /^[0-9]{4}[－\- .～][0-9]{1}/, f: x => `${x.substr(0, 4)}年0${x[5]}月` },
  { id: 'nomonth', re: /^[0-9]{4}/, f: x => NOPES },
  { id: 'bad', re: /^202年6月$/, f: x => NOPES }
]
let YearMonthDay = [
  { id: 'good', re: /^\d{4}-\d{2}-\d{2}$/, f: x => x },
  { id: 'dots_or_wronglen', 
      re: /^\d{4}(\-|\.|－|--|—|年| |～|- |一|_| -|---|―|\/)\d{1,2}(\-|\.|－|--|—|月| |～|- |一|_| -|日|---|―|----|\/)\d{1,2}[日号曰]?$/,
      f: x => {
        // console.log(x)
        x = x.replace('--', '-').replace('--', '-').replace('--', '-').replace('--', '-').
        replace(/(\.|－|--|—|年| |～|- |一|_| -|―|\/)/, '-').
        replace(/(\.|－|--|—|月| |～|- |一|_| -|日|―|\/)/, '-').
        replace(/[日号曰]/, '').split('-')
        try {
          if (x[1].length === 1) x[1] = '0' + x[1]
          if (x[2].length === 1) x[2] = '0' + x[2]
        } catch(e) {
          console.log('error: ', e)
          console.log(x)
        }
        // console.log(x.join('-'))
        return x.join('-')
      }
  },
  { id: 'midhyphen', re:/^\d{4}[\-\.]\d{4}(-20220609)?$/, f: x => `${x.substr(0, 4)}-${x.substr(5, 2)}-${x.substr(7, 2)}` },
  { id: 'full_len', re: /^[0-9]{8}(-02-15)?$/, f: x => `${x.substr(0, 4)}-${x.substr(4, 2)}-${x.substr(6, 2)}` },
  { id: 'full_len_weird', re: /^[0-9]{7}$/, f: x => {
    console.log('weird', x)
    // mdd
    let mdd = parseInt(x.substr(5, 2)) <= 31
    let mmd = parseInt(x.substr(4, 2)) <= 12
    if (mdd === mmd) {
      console.log('weird', x)
      return NOPES
    } else {
      return `${x.substr(0, 4)}-0${x.substr(4, mdd ? 1 : 2)}-${x.substr(5, mdd ? 2 : 1)}`
    }
  } },
  { id: '6digits', re: /^[0-9]{6}$/, f: x => NOPES },
  { id: '无', re: /^(无)$/, f: x => NOPES },
  { id: 'preserve', re: /已过世|无限/, f: x => x },
  { id: 'noIDCard', re: /暂没办理|无身份证|未办理身份证|暂无|\//, f: x => '无身份证' },
  { id: 'wtf', re: /20/, f: x => NOPES },
  { id: 'no日', re: /^\d{4}年\d{1,2}月$/, f: x => NOPES },
  { id: 'no年', re: /^\d{1,2}月\d{1,2}日$/, f: x => NOPES },
  { id: 'IDNum', re: /^\d{17}[\dX]$/, f: x => {
    x = x.substr(6, 8)
    return `${x.substr(0, 4)}-${x.substr(4, 2)}-${x.substr(6, 2)}`
  } }
]
let Patterns = {
  '组件 民族': EthiPattern,
  '组件 母亲民族': EthiPattern,
  '组件 父亲民族': EthiPattern,
  '组件 入团年月': YearMonth,
  '组件 出生日期': YearMonthDay,
  '组件 身份证件有效期至': YearMonthDay,
  '组件 身份证件生效日期': YearMonthDay,
  '组件 母亲出生日期': YearMonthDay,
  '组件 父亲出生日期': YearMonthDay,
  '组件 初中入学年月': YearMonth,
  '组件 初中毕业年月': YearMonth,
  '组件 小学入学年月': YearMonth,
  '组件 小学毕业年月': YearMonth
}

function matchPatterns(lp, ps) {
  let counter = { fail: 0 }
  let shown = false
  for (let i = 1; i < lp.length; i++) {
    let r = match(lp[i], ps)
    if (!counter[r.id]) {
      counter[r.id] = 0
    }
    counter[r.id] += 1
    lp[i] = r.res
    if (r.id === 'fail' && !shown) {
      console.log('fail at', lp[i], 'No.', i)
      shown = true
    }
  }
  console.log(counter)
  if (counter.fail > 0) console.log('================================')
}

function match(s, ps) {
  if (!s) {
    return { id: 'empty', res: NOPES }
  }
  for (let i = 0; i < ps.length; i++) {
    const p = ps[i]
    if (p.re.exec(s)) {
      return { id: p.id, res: p.f(s) }
    }
  }
  return { id: 'fail', res: s }
}

function transpose(mat) {
  return mat[0].map(function(col, i) {
    return mat.map(function(row) {
      return row[i]
    })
  })
}

async function main() {
  let sheets = xlsx.parse('hand.xlsx')
  // console.log(sheets)
  await sheets.forEach(sheet => {
    let neat = transpose(sheet.data)
    neat.forEach(item => {
      if (Patterns[item[0]]) {
        let pattern = Patterns[item[0]]
        console.log('Processing #' + item[0])
        matchPatterns(item, pattern)
        matchPatterns(item, pattern)
      } else {
        console.log('Please consider #' + item[0])
      }
    })
    sheet.data = transpose(neat)
  })
  let buffer = await xlsx.build(sheets)
  fs.writeFileSync('res.xlsx', buffer)
}

main()

