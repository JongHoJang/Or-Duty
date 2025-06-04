'use client'

import { useState } from 'react'
import * as XLSX from 'sheetjs-style'

interface DutyEntry {
  name: string
  isOrange: boolean
}

interface DutyMapByCode {
  originalKeys: string[]
  entries: DutyEntry[]
}

type DutyMap = Record<number, Record<string, DutyMapByCode>>

const customOrder = [
  ...Array.from({ length: 16 }, (_, i) => String(i + 1)),
  'D5',
  'DD',
  'DD12',
  'DB',
  'DC',
  'C',
  'P3',
  'EA',
  'E1',
  'N',
  'NN12',
  'OFF',
  'ST',
]

function getDutyPriority(code: string): [number, number | string] {
  const upper = code.toUpperCase()
  const index = customOrder.indexOf(upper)
  if (index !== -1) return [0, index]
  return [1, upper]
}

function normalizeDutyKey(code: string): string {
  if (typeof code !== 'string') code = String(code ?? '')
  const upper = code.toUpperCase().replace(/[^A-Z0-9]/g, '')
  if (/^NN12$/.test(upper)) return 'NN12'
  if (/^V$|^VF$|^OV$|^W$/.test(upper)) return 'OFF'
  if (/^OF$/.test(upper)) return 'OFF'
  if (/^P$/.test(upper)) return 'P3'
  if (/^C$/.test(upper)) return 'C'
  if (/^E1$/.test(upper)) return 'E1'
  if (/^ST$/.test(upper)) return 'ST'
  if (/^D5$/.test(upper)) return 'D5'
  if (/^DB$/.test(upper)) return 'DB'
  if (/^DC$/.test(upper)) return 'DC'
  if (/^\d+$/.test(upper)) return upper
  if (upper.startsWith('D')) return upper
  if (upper.startsWith('E')) return upper
  if (upper.startsWith('N')) return 'N'
  return upper
}

function getDisplayLabel(normKey: string): string {
  if (/^\d+$/.test(normKey)) return normKey + 'R'
  if (normKey === 'C') return '세포'
  if (normKey === 'E1') return 'E'
  if (normKey === 'OFF') return 'OFF'
  if (normKey === 'ST') return 'Station'
  return normKey
}

export default function UploadPage() {
  const [excelData, setExcelData] = useState<string[][]>([])
  const [resultByDay, setResultByDay] = useState<DutyMap>({})
  const [uploaded, setUploaded] = useState(false)

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = e => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const workbook = XLSX.read(data, { type: 'array', cellStyles: true })
      const sheetName = workbook.SheetNames[workbook.SheetNames.length - 1]
      const worksheet = workbook.Sheets[sheetName]
      const raw = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      }) as unknown as string[][]

      const dateRow = raw[0]
      // const dayRow = raw[1] || []
      const contentRows = raw.slice(2)
      const byDay: DutyMap = {}

      dateRow.forEach((date, colIndex) => {
        if (typeof date === 'number' && colIndex >= 2) {
          const dutyMap: Record<
            string,
            {
              originalKeys: string[]
              entries: { name: string; isOrange: boolean }[]
            }
          > = {}

          contentRows.forEach((row, rowIndex) => {
            const name = row[1]
            const duty = row[colIndex]
            const cellAddress = XLSX.utils.encode_cell({
              r: rowIndex + 2,
              c: colIndex,
            })
            const cell = worksheet[cellAddress]
            const fgRgb = cell?.s?.fgColor?.rgb?.toUpperCase()
            const isOrange = fgRgb === 'FFC000'

            if (duty) {
              const normalized = normalizeDutyKey(duty)
              if (!dutyMap[normalized]) {
                dutyMap[normalized] = { originalKeys: [duty], entries: [] }
              } else if (!dutyMap[normalized].originalKeys.includes(duty)) {
                dutyMap[normalized].originalKeys.push(duty)
              }
              dutyMap[normalized].entries.push({ name, isOrange })
            }
          })

          for (const key in dutyMap) {
            dutyMap[key].entries.sort(
              (a, b) =>
                contentRows.findIndex(r => r[1] === a.name) -
                contentRows.findIndex(r => r[1] === b.name)
            )
          }

          if (Object.keys(dutyMap).length > 0) {
            byDay[date] = dutyMap
          }
        }
      })

      setExcelData(raw)
      setResultByDay(byDay)
      setUploaded(true)
    }

    reader.readAsArrayBuffer(file)
  }

  return (
    <div className="p-6 max-w-6xl mx-auto">
      {!uploaded && (
        <div className="min-h-screen flex flex-col items-center justify-center text-center pb-40">
          <div className="bg-gray-100 p-10 rounded-xl shadow-xl">
            <div className="flex flex-row items-end mb-4">
              <h1 className="text-2xl font-bold mb-4">
                듀티표를 업로드 해주세요.
              </h1>
            </div>
            <label className="inline-block px-4 py-2 bg-blue-500 text-white rounded cursor-pointer hover:bg-blue-600">
              파일 선택
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
          <a
            href="https://walnut-hose-a93.notion.site/208ebd82d491809d921fe226f3a6ddba?source=copy_link"
            target="_blank"
            rel="noopener noreferrer"
            className="text-blue-700 underline cursor-pointer mt-4"
          >
            <span>이용 가이드 보러가기 &gt;</span>
          </a>
        </div>
      )}

      <div className="grid grid-cols-5 gap-4">
        {Object.entries(resultByDay).map(([day, duties]) => {
          const sortedDutyEntries = Object.entries(duties).sort((a, b) => {
            const [aPriority, aSub] = getDutyPriority(a[0])
            const [bPriority, bSub] = getDutyPriority(b[0])
            if (aPriority !== bPriority) return aPriority - bPriority
            if (typeof aSub === 'number' && typeof bSub === 'number')
              return aSub - bSub
            return aSub.toString().localeCompare(bSub.toString())
          })

          const weekday = excelData[1]?.[Number(day) + 1] || ''
          const numericDutyKeys = Array.from({ length: 16 }, (_, i) =>
            String(i + 1)
          )
          const numericKeysSet = new Set(numericDutyKeys)
          const numericCount = sortedDutyEntries.filter(([key]) =>
            numericKeysSet.has(key)
          ).length

          return (
            <div key={day} className="h-[180mm]">
              <h2 className="font-semibold text-md mb-1 text-center border p-1 rounded-sm">
                {day}일 {weekday && `(${weekday})`}
              </h2>
              <div className="border rounded-sm p-2 bg-white flex-1 overflow-hidden">
                <div className="grid grid-cols-1 gap-x-6 gap-y-1 overflow-hidden">
                  {sortedDutyEntries.map(([normKey, { entries }], index) => {
                    const isLastNumeric = index === numericCount - 1
                    return (
                      <div key={normKey}>
                        <div className="flex flex-row items-start">
                          <h3 className="font-semibold text-xs w-[40px]">
                            {getDisplayLabel(normKey)}
                          </h3>
                          <ul className="list-none text-xs grid-cols-3 grid ml-1">
                            {entries.map((entry, idx) => (
                              <li
                                key={idx}
                                className={`w-12 px-1 rounded text-center ${
                                  entry.isOrange ? 'bg-gray-300 font-bold' : ''
                                }`}
                              >
                                {entry.name}
                              </li>
                            ))}
                          </ul>
                        </div>
                        {isLastNumeric && numericCount > 0 && (
                          <hr className="my-2" />
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}
