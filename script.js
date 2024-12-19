const fileInputCustomer = document.getElementById('fileInputCustomer')
const fileInputOwn = document.getElementById('fileInputOwn')
const downloadJsonButton = document.getElementById('downloadJsonButton')
const downloadExcelButton = document.getElementById('downloadExcelButton')
const analyzeButton = document.getElementById('analyzeButton')
const displayArea = document.querySelector('.display-area')
const downloadResultButton = document.getElementById('downloadResultButton')

let customerData = null
let ownData = null

// Function to parse Excel file to JSON
function parseExcelToJson(file, fileType) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result)
      const workbook = XLSX.read(data, { type: 'array' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      // Get the range of cells
      const range = XLSX.utils.decode_range(worksheet['!ref'])
      const headers = []
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: range.s.r, c: C })
        const cell = worksheet[address]
        headers.push(cell ? cell.v : `Column ${C}`) // Use 'Column C' if no header
      }
      console.log('Headers:', headers) // Log the headers

      let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
      // Convert array of arrays to array of objects
      jsonData = jsonData.slice(1).map((row) => {
        const obj = {}
        headers.forEach((header, index) => {
          obj[header] = row[index] || ''
        })
        if (fileType === 'customer') {
          const documentNumber = obj['Document Number'] || ''
          const lineId = obj['Line ID'] || ''
          obj['Document Number-Line ID'] =
            documentNumber && lineId ? `${documentNumber}-${lineId}` : ''
        } else if (fileType === 'own') {
          const po = obj['Po'] || ''
          const line = obj['line'] || ''
          console.log('PO value:', po, 'Line value:', line) // Log values before concatenation
          obj['PO-Line'] = po && line ? `${po}-${line}` : ''
        }
        return obj
      })
      resolve(jsonData)
    }
    reader.onerror = (error) => {
      reject(error)
    }
    reader.readAsArrayBuffer(file)
  })
}

// Function to handle customer file upload
fileInputCustomer.addEventListener('change', async (event) => {
  const file = event.target.files[0]
  if (file) {
    try {
      customerData = await parseExcelToJson(file, 'customer')
      console.log('Customer data loaded:', customerData)
    } catch (err) {
      console.log('error while loading customer file', err)
    }
  }
})

// Function to handle own file upload
fileInputOwn.addEventListener('change', async (event) => {
  const file = event.target.files[0]
  if (file) {
    try {
      ownData = await parseExcelToJson(file, 'own')
      console.log('Own data loaded:', ownData)
    } catch (err) {
      console.log('error while loading your file', err)
    }
  }
})

// Function to download the data as JSON files
downloadJsonButton.addEventListener('click', () => {
  if (customerData) {
    downloadJson(customerData, 'customer_data.json')
  }
  if (ownData) {
    downloadJson(ownData, 'own_data.json')
  }
})

function downloadJson(jsonData, filename) {
  const json = JSON.stringify(jsonData, null, 2)
  const blob = new Blob([json], { type: 'application/json' })
  const url = URL.createObjectURL(blob)

  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

// Function to download the excel files from json
downloadExcelButton.addEventListener('click', () => {
  if (customerData) {
    downloadExcel(customerData, 'customer_data.xlsx')
  }
  if (ownData) {
    downloadExcel(ownData, 'own_data.xlsx')
  }
})

function downloadExcel(jsonData, filename) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new()

  // Convert JSON to sheet
  const worksheet = XLSX.utils.json_to_sheet(jsonData)

  // Append the sheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

  // Write the workbook to a binary string
  const excelBinary = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' })

  // Convert the binary string to a blob
  const blob = new Blob([s2ab(excelBinary)], {
    type: 'application/octet-stream',
  })

  // Create a URL for the blob
  const url = URL.createObjectURL(blob)

  // Create a link to download the file
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

function s2ab(s) {
  const buf = new ArrayBuffer(s.length)
  const view = new Uint8Array(buf)
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xff
  }
  return buf
}

// Function to compare the data and display the results
analyzeButton.addEventListener('click', () => {
  if (!customerData || !ownData) {
    displayArea.innerHTML =
      '<p class="error-message">Please upload both customer and your excel files.</p>'
    return
  }

  // Clear the display area
  displayArea.innerHTML = ''

  // Create the timer display
  const timerDisplay = document.createElement('div')
  timerDisplay.id = 'timer-display'
  timerDisplay.style.display = 'none' // Initially hidden
  displayArea.appendChild(timerDisplay)

  const table = document.createElement('table')
  table.classList.add('comparison-table')
  const thead = document.createElement('thead')

  // --- Create Alphabetical Header Row ---
  const alphabetHeaderRow = document.createElement('tr')
  const customerHeadersCount = Object.keys(customerData[0]).length
  const totalHeadersCount = customerHeadersCount + 1 + 5 + 1 // +1 for 'Match Status', +5 for added headers, +1 for Serial Number

  // Add empty cell for the serial number column in the alphabetical header row
  alphabetHeaderRow.appendChild(document.createElement('th'))

  for (let i = 0; i < totalHeadersCount - 1; i++) {
    const th = document.createElement('th')
    th.textContent = String.fromCharCode(65 + i) // 65 is ASCII for 'A'
    alphabetHeaderRow.appendChild(th)
  }
  thead.appendChild(alphabetHeaderRow)
  // --- End of Alphabetical Header Row ---

  // --- Create Original Header Row ---
  const headerRow = document.createElement('tr')

  // Add header for the serial number column
  const serialHeader = document.createElement('th')
  serialHeader.textContent = 'Sr. No.'
  headerRow.appendChild(serialHeader)

  const customerHeaders = Object.keys(customerData[0])
  customerHeaders.forEach((header) => {
    const th = document.createElement('th')
    th.textContent = header
    headerRow.appendChild(th)
  })
  const statusHeader = document.createElement('th')
  statusHeader.textContent = 'Match Status'
  headerRow.appendChild(statusHeader)

  const ownHeaders = [
    'BD xf-date',
    'Prouction status(SFC210)',
    'Order QTY',
    'Shipmod',
    'unit price',
  ]
  ownHeaders.forEach((header) => {
    const th = document.createElement('th')
    th.textContent = header
    headerRow.appendChild(th)
  })

  thead.appendChild(headerRow)
  // --- End of Original Header Row ---

  table.appendChild(thead)

  const tbody = document.createElement('tbody')
  const styledData = []
  let serialNumber = 1 // Initialize serial number

  customerData.forEach((customerRow) => {
    const tr = document.createElement('tr')
    const customerId = customerRow['Document Number-Line ID']
    let matchFound = false
    let matchedOwnRow = null
    const styledRow = {}
    for (const ownRow of ownData) {
      if (ownRow['PO-Line'] === customerId) {
        matchFound = true
        matchedOwnRow = ownRow
        break
      }
    }

    // Add serial number cell
    const serialNumberCell = document.createElement('td')
    serialNumberCell.textContent = serialNumber++ // Set serial number and increment
    tr.appendChild(serialNumberCell)
    styledRow['Sr. No.'] = serialNumberCell.textContent

    customerHeaders.forEach((header) => {
      const td = document.createElement('td')
      td.textContent = customerRow[header]
      tr.appendChild(td)
      styledRow[header] = customerRow[header]
    })

    const statusTd = document.createElement('td')
    if (matchFound) {
      statusTd.textContent = 'Match'
      statusTd.classList.add('match')
    } else {
      statusTd.textContent = 'No Match'
      statusTd.classList.add('mismatch')
    }
    tr.appendChild(statusTd)
    styledRow['Match Status'] = statusTd.textContent

    // Add own data
    ownHeaders.forEach((header) => {
      const td = document.createElement('td')
      let ownValue = ''
      if (matchedOwnRow) {
        const customerValue = customerRow[header] || ''
        ownValue = matchedOwnRow[header] || ''
        td.textContent = ownValue
        if (customerValue != ownValue) {
          td.classList.add('mismatch-cell')
        }
      } else {
        td.textContent = ''
      }
      tr.appendChild(td)
      styledRow[header] = ownValue
    })

    tbody.appendChild(tr)
    styledData.push(styledRow)
  })

  table.appendChild(tbody)
  // --- Analyze Column W and N ---
  const columnNIndex = Object.keys(styledData[0]).indexOf(
    Object.keys(customerData[0])[12]
  ) // Assuming N is the 14th column (index 13) in customerData, so find the index of 13th element in styledData
  const columnWIndex = Object.keys(styledData[0]).indexOf(ownHeaders[4]) // Assuming W is the last column (index 22)

  for (let i = 0; i < styledData.length; i++) {
    const rowData = styledData[i]
    const columnWValue = rowData[Object.keys(rowData)[columnWIndex]] || ''
    const columnNValue = rowData[Object.keys(rowData)[columnNIndex]] || ''

    if (columnWValue !== columnNValue) {
      // Highlight the cells in the table
      const rowIndex = i + 2 // +2 to account for the two header rows
      const tableRow = table.querySelector(`tbody tr:nth-child(${rowIndex})`)
      if (tableRow) {
        const cellW = tableRow.querySelector(
          `td:nth-child(${columnWIndex + 2})`
        )
        const cellN = tableRow.querySelector(
          `td:nth-child(${columnNIndex + 2})`
        )

        if (cellW && cellN) {
          cellW.classList.add('mismatch-highlight')
          cellN.classList.add('mismatch-highlight')
        } else {
          console.error('Could not find cellW or cellN', cellW, cellN)
        }
      } else {
        console.error('Could not find tableRow for rowIndex:', rowIndex)
      }

      // Highlight the cells in styled data
      styledData[i][Object.keys(rowData)[columnWIndex]] = {
        value: columnWValue,
        style: { fill: { fgColor: { rgb: 'FFCCCC' } } },
      }
      styledData[i][Object.keys(rowData)[columnNIndex]] = {
        value: columnNValue,
        style: { fill: { fgColor: { rgb: 'FFCCCC' } } },
      }
    }
  }
  // --- End of Column W and N Analysis ---

  // --- Hide table and display message ---
  // Append the table to the display area
  displayArea.appendChild(table)
  // Hide the table
  table.style.display = 'none'

  // --- Timer Logic ---
  let timeLeft = 5 // 5 seconds countdown
  timerDisplay.textContent = `Time left: ${timeLeft}s`
  timerDisplay.style.display = 'block' // Show the timer

  const timerInterval = setInterval(() => {
    timeLeft--
    timerDisplay.textContent = `Time left: ${timeLeft}s`

    if (timeLeft <= 0) {
      clearInterval(timerInterval)
      timerDisplay.style.display = 'none' // Hide the timer

      // Create and display the message
      const messageContainer = document.createElement('div')
      messageContainer.id = 'message-container'
      messageContainer.textContent = 'Analyze completed and ready to download'
      messageContainer.classList.add('show-message') // Add the class to trigger the animation
      displayArea.appendChild(messageContainer)
    }
  }, 1000) // Update every 1 second (1000 milliseconds)
  // --- End of Timer Logic ---

  downloadResultButton.styledData = styledData
})

// Function to download the result as excel
downloadResultButton.addEventListener('click', () => {
  const table = document.querySelector('.comparison-table')
  if (!table) {
    displayArea.innerHTML =
      '<p class="error-message">Please click on analyze button first.</p>'
    return
  }
  const jsonData = tableToJson(table)
  downloadExcelWithHighlight(
    downloadResultButton.styledData,
    'comparison_result.xlsx'
  )
})

function tableToJson(table) {
  const headers = []
  const data = []

  // Get headers (skip the first column which is Sr.No.)
  const headerRow = table.querySelector('thead tr:nth-child(2)')
  for (let i = 0; i < headerRow.cells.length; i++) {
    headers.push(headerRow.cells[i].textContent)
  }

  // Get table data (add serial number to each row)
  const trs = table.querySelectorAll('tbody tr')
  trs.forEach((tr, index) => {
    const rowData = {}
    // Add serial number to row data
    rowData['Sr. No.'] = index + 1
    const tds = tr.querySelectorAll('td')
    tds.forEach((td, index) => {
      // Offset index by 1 to account for the serial number column
      rowData[headers[index]] = td.textContent
    })
    data.push(rowData)
  })

  return data
}

function downloadExcelWithHighlight(jsonData, filename) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new()

  // Convert JSON to sheet
  const worksheet = XLSX.utils.json_to_sheet(jsonData)

  // Apply highlighting
  const range = XLSX.utils.decode_range(worksheet['!ref'])
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
      const cell = worksheet[cellAddress]
      if (
        cell &&
        jsonData[R - 1] &&
        jsonData[R - 1][Object.keys(jsonData[R - 1])[C]]
      ) {
        const header = Object.keys(jsonData[R - 1])[C]
        const customerValue = jsonData[R - 1][header] || ''
        const ownValue = jsonData[R - 1][header] || ''

        // Check if the cell has highlighting from styledData
        if (
          jsonData[R - 1][header] &&
          typeof jsonData[R - 1][header] === 'object' &&
          jsonData[R - 1][header].style
        ) {
          worksheet[cellAddress].s = jsonData[R - 1][header].style
        }

        if (header != 'Match Status' && customerValue != ownValue) {
          worksheet[cellAddress].s = { fill: { fgColor: { rgb: 'FFCCCC' } } }
        }
      }
    }
  }

  // Append the sheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

  // Write the workbook to a binary string
  const excelBinary = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' })

  // Convert the binary string to a blob
  const blob = new Blob([s2ab(excelBinary)], {
    type: 'application/octet-stream',
  })

  // Create a URL for the blob
  const url = URL.createObjectURL(blob)

  // Create a link to download the file
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
