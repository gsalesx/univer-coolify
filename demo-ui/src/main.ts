import './style.css'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

import { createDefaultWorkbookData, setupUniver } from './setup-univer'

type WorkbookListItem = {
  id: string
  name: string
  createdAt: string
  updatedAt: string
}

type WorkbookResponse = {
  workbook: WorkbookListItem & {
    snapshot: Record<string, unknown>
  }
}

type ImportedXlsxImage = {
  row: number
  column: number
  blob: Blob
  fileName: string
}

const STATUS_COLUMN_INDEX = 7
const STATUS_COLUMN_NAME = 'Status'
const STATUS_OPTIONS = [
  { label: 'Pronto', color: '#c084fc' },
  { label: 'Separado', color: '#93c5fd' },
  { label: 'Manual', color: '#fdba74' },
  { label: 'Editar', color: '#fca5a5' },
  { label: 'Cancelado', color: '#dc2626' },
  { label: 'Aprovado', color: '#86efac' },
  { label: 'Sem fotos', color: '#d1d5db' },
]

let selectedUploadFile: File | null = null

function setupQuickToolbar(univerAPI: ReturnType<typeof setupUniver>) {
  document.querySelectorAll<HTMLButtonElement>('.quick-color').forEach((button) => {
    button.addEventListener('click', () => {
      const color = button.dataset.fillColor
      const activeRange = univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSelection()?.getActiveRange()

      if (!color || !activeRange) return

      activeRange.setBackgroundColor(color)
      activeRange.setBorder(univerAPI.Enum.BorderType.ALL, univerAPI.Enum.BorderStyleTypes.THIN, '#d1d5db')
    })
  })
}

function workbookUrl(id: string) {
  return `/sheet?id=${encodeURIComponent(id)}`
}

function workbookNameFromFile(file: File) {
  return file.name.replace(/\.(xlsx|xls)$/i, '') || 'Imported spreadsheet'
}

function normalizeWorkbookName(value: string) {
  return value.trim() || 'Untitled spreadsheet'
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat('pt-BR', {
    dateStyle: 'short',
    timeStyle: 'short',
  }).format(new Date(value))
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}

function setSelectedUploadFile(file: File | null) {
  selectedUploadFile = file

  const fileName = document.querySelector<HTMLElement>('#upload-file-name')
  const nameInput = document.querySelector<HTMLInputElement>('#upload-workbook-name')
  const submitButton = document.querySelector<HTMLButtonElement>('#submit-upload')

  if (fileName) fileName.textContent = file ? file.name : ''
  if (file && nameInput && !nameInput.value.trim()) nameInput.value = workbookNameFromFile(file)
  if (submitButton) submitButton.disabled = !file
}

async function saveImportedWorkbook(snapshot: Record<string, unknown>, name: string) {
  const response = await fetch('/api/workbooks/import', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, snapshot }),
  })

  if (!response.ok) throw new Error('Nao foi possivel salvar a planilha importada.')

  return (await response.json()) as { workbook: WorkbookListItem }
}

async function renameWorkbook(id: string, name: string) {
  const response = await fetch(`/api/workbooks/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
  })

  if (!response.ok) throw new Error('Nao foi possivel renomear a planilha.')
}

function getElementsByLocalName(element: ParentNode, name: string) {
  return Array.from(element.querySelectorAll('*')).filter((node) => node.localName === name)
}

function getFirstChildText(element: Element, name: string) {
  return getElementsByLocalName(element, name)[0]?.textContent || ''
}

function getRelationshipId(element: Element) {
  const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

  return element.getAttribute('r:id')
    || element.getAttribute('r:embed')
    || element.getAttributeNS(relationshipNamespace, 'id')
    || element.getAttributeNS(relationshipNamespace, 'embed')
    || ''
}

function resolveZipPath(fromPart: string, target: string) {
  if (target.startsWith('/')) return target.slice(1)

  const stack = fromPart.split('/').slice(0, -1)

  target.split('/').forEach((part) => {
    if (!part || part === '.') return
    if (part === '..') stack.pop()
    else stack.push(part)
  })

  return stack.join('/')
}

function relationshipPathForPart(partPath: string) {
  const parts = partPath.split('/')
  const fileName = parts.pop()

  return `${parts.join('/')}/_rels/${fileName}.rels`
}

async function readRelationships(zip: JSZip, partPath: string) {
  const relationships = new Map<string, string>()
  const relsFile = zip.file(relationshipPathForPart(partPath))

  if (!relsFile) return relationships

  const document = new DOMParser().parseFromString(await relsFile.async('string'), 'application/xml')

  getElementsByLocalName(document, 'Relationship').forEach((relationship) => {
    const id = relationship.getAttribute('Id')
    const target = relationship.getAttribute('Target')

    if (id && target) relationships.set(id, resolveZipPath(partPath, target))
  })

  return relationships
}

function mimeTypeFromPath(path: string) {
  const extension = path.split('.').pop()?.toLowerCase()

  if (extension === 'png') return 'image/png'
  if (extension === 'gif') return 'image/gif'
  if (extension === 'webp') return 'image/webp'
  return 'image/jpeg'
}

async function getWorksheetPartsByName(zip: JSZip) {
  const workbookPart = 'xl/workbook.xml'
  const workbookFile = zip.file(workbookPart)
  const worksheets = new Map<string, string>()

  if (!workbookFile) return worksheets

  const workbookRelationships = await readRelationships(zip, workbookPart)
  const document = new DOMParser().parseFromString(await workbookFile.async('string'), 'application/xml')

  getElementsByLocalName(document, 'sheet').forEach((sheet) => {
    const name = sheet.getAttribute('name')
    const relationshipId = getRelationshipId(sheet)
    const worksheetPart = workbookRelationships.get(relationshipId)

    if (name && worksheetPart) worksheets.set(name, worksheetPart)
  })

  return worksheets
}

async function extractXlsxImages(file: File, fileBuffer: ArrayBuffer) {
  if (!file.name.toLowerCase().endsWith('.xlsx')) return new Map<string, ImportedXlsxImage[]>()

  const zip = await JSZip.loadAsync(fileBuffer)
  const worksheets = await getWorksheetPartsByName(zip)
  const imagesBySheet = new Map<string, ImportedXlsxImage[]>()

  for (const [sheetName, worksheetPart] of worksheets) {
    const worksheetFile = zip.file(worksheetPart)
    if (!worksheetFile) continue

    const worksheetRelationships = await readRelationships(zip, worksheetPart)
    const worksheetDocument = new DOMParser().parseFromString(await worksheetFile.async('string'), 'application/xml')

    for (const drawing of getElementsByLocalName(worksheetDocument, 'drawing')) {
      const drawingPart = worksheetRelationships.get(getRelationshipId(drawing))
      const drawingFile = drawingPart ? zip.file(drawingPart) : null

      if (!drawingPart || !drawingFile) continue

      const drawingRelationships = await readRelationships(zip, drawingPart)
      const drawingDocument = new DOMParser().parseFromString(await drawingFile.async('string'), 'application/xml')
      const anchors = [
        ...getElementsByLocalName(drawingDocument, 'twoCellAnchor'),
        ...getElementsByLocalName(drawingDocument, 'oneCellAnchor'),
      ]

      for (const anchor of anchors) {
        const blip = getElementsByLocalName(anchor, 'blip')[0]
        const mediaPath = blip ? drawingRelationships.get(getRelationshipId(blip)) : null
        const mediaFile = mediaPath ? zip.file(mediaPath) : null

        if (!mediaPath || !mediaFile) continue

        const marker = getElementsByLocalName(anchor, 'from')[0]
        const row = Number(getFirstChildText(marker, 'row'))
        const column = Number(getFirstChildText(marker, 'col'))

        if (!Number.isFinite(row) || !Number.isFinite(column)) continue

        const blob = await mediaFile.async('blob')
        const images = imagesBySheet.get(sheetName) || []

        images.push({
          row,
          column,
          blob: new Blob([blob], { type: mimeTypeFromPath(mediaPath) }),
          fileName: mediaPath.split('/').pop() || 'image.jpg',
        })
        imagesBySheet.set(sheetName, images)
      }
    }
  }

  return imagesBySheet
}

function escapeFormulaString(value: string) {
  return value.replace(/"/g, '""')
}

async function uploadImportedImage(image: ImportedXlsxImage) {
  const body = new FormData()
  body.append('image', image.blob, image.fileName)

  const response = await fetch('/api/images/upload', {
    method: 'POST',
    body,
  })

  if (!response.ok) throw new Error('Nao foi possivel salvar uma imagem importada.')

  return (await response.json()) as { url: string }
}

async function importXlsFile(file: File) {
  const fileBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: 'array' })
  const imagesBySheet = await extractXlsxImages(file, fileBuffer)
  const snapshot = createDefaultWorkbookData('imported-workbook', workbookNameFromFile(file)) as any

  snapshot.sheetOrder = []
  snapshot.sheets = {}

  for (const [index, sheetName] of workbook.SheetNames.entries()) {
    const worksheet = workbook.Sheets[sheetName]
    const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : null
    const sheetId = `sheet-${index + 1}`
    let rowCount = range ? Math.max(range.e.r + 1, 100) : 100
    const sourceColumnCount = range ? range.e.c + 1 : 0
    let columnCount = Math.max(sourceColumnCount >= STATUS_COLUMN_INDEX + 1 ? sourceColumnCount + 1 : STATUS_COLUMN_INDEX + 1, 20)
    const cellData: Record<string, Record<string, Record<string, unknown>>> = {}
    const maxColumnTextLengths: Record<number, number> = {}
    const rowData: Record<number, { h: number }> = {}

    if (range) {
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let column = range.s.c; column <= range.e.c; column += 1) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: column })]
          if (!cell) continue

          const targetColumn = column >= STATUS_COLUMN_INDEX ? column + 1 : column
          const value = cell.w ?? cell.v ?? ''
          const textLength = String(value).split(/\r?\n/).reduce((max, line) => Math.max(max, line.length), 0)

          maxColumnTextLengths[targetColumn] = Math.max(maxColumnTextLengths[targetColumn] || 0, textLength)
          cellData[row] ||= {}
          cellData[row][targetColumn] = {
            v: value,
            ...(cell.f ? { f: cell.f.startsWith('=') ? cell.f : `=${cell.f}` } : {}),
          }
        }
      }
    }

    for (const image of imagesBySheet.get(sheetName) || []) {
      const targetColumn = image.column >= STATUS_COLUMN_INDEX ? image.column + 1 : image.column
      const uploadedImage = await uploadImportedImage(image)

      cellData[image.row] ||= {}
      cellData[image.row][targetColumn] = {
        v: '',
        f: `=IMAGE("${escapeFormulaString(uploadedImage.url)}","Foto",1)`,
      }
      rowCount = Math.max(rowCount, image.row + 1)
      columnCount = Math.max(columnCount, targetColumn + 1)
      rowData[image.row] = { h: Math.max(rowData[image.row]?.h || 0, 96) }
      maxColumnTextLengths[targetColumn] = Math.max(maxColumnTextLengths[targetColumn] || 0, 14)
    }

    cellData[0] ||= {}
    cellData[0][STATUS_COLUMN_INDEX] = {
      v: STATUS_COLUMN_NAME,
      s: {
        bl: 1,
        bg: { rgb: '#f3f4f6' },
      },
    }
    maxColumnTextLengths[STATUS_COLUMN_INDEX] = Math.max(maxColumnTextLengths[STATUS_COLUMN_INDEX] || 0, STATUS_COLUMN_NAME.length + 4)

    const columnData = Object.fromEntries(
      Object.entries(maxColumnTextLengths).map(([column, length]) => {
        const width = Math.min(Math.max(length * 8 + 28, 93), 420)
        return [column, { w: width }]
      }),
    )

    snapshot.sheetOrder.push(sheetId)
    snapshot.sheets[sheetId] = {
      ...createDefaultWorkbookData(sheetId, sheetName).sheets['sheet-01'],
      id: sheetId,
      name: sheetName || `Sheet${index + 1}`,
      rowCount,
      columnCount,
      cellData,
      columnData,
      rowData,
    }
  }

  if (snapshot.sheetOrder.length === 0) {
    const sheetId = 'sheet-1'
    snapshot.sheetOrder = [sheetId]
    snapshot.sheets[sheetId] = createDefaultWorkbookData(sheetId, 'Sheet1').sheets['sheet-01']
  }

  return snapshot
}

async function uploadSelectedWorkbook() {
  if (!selectedUploadFile) return

  const submitButton = document.querySelector<HTMLButtonElement>('#submit-upload')
  const file = selectedUploadFile
  const extension = file.name.split('.').pop()?.toLowerCase()

  if (extension !== 'xlsx' && extension !== 'xls') {
    alert('Envie um arquivo .xls ou .xlsx.')
    return
  }

  if (submitButton) {
    submitButton.disabled = true
    submitButton.textContent = 'Importando...'
  }

  try {
    const snapshot = await importXlsFile(file)
    const nameInput = document.querySelector<HTMLInputElement>('#upload-workbook-name')
    const data = await saveImportedWorkbook(snapshot, normalizeWorkbookName(nameInput?.value || workbookNameFromFile(file)))
    window.location.href = workbookUrl(data.workbook.id)
  } catch (error) {
    console.error(error)
    alert('Nao foi possivel importar a planilha agora.')
  } finally {
    if (submitButton) {
      submitButton.disabled = !selectedUploadFile
      submitButton.textContent = 'Upload'
    }
  }
}

function closeUploadModal() {
  const modal = document.querySelector<HTMLElement>('#upload-modal')
  const input = document.querySelector<HTMLInputElement>('#workbook-upload-input')
  const nameInput = document.querySelector<HTMLInputElement>('#upload-workbook-name')

  if (modal) modal.hidden = true
  if (input) input.value = ''
  if (nameInput) nameInput.value = ''
  setSelectedUploadFile(null)
}

function setupUploadModal() {
  const modal = document.querySelector<HTMLElement>('#upload-modal')
  const openButton = document.querySelector<HTMLButtonElement>('#open-upload-modal')
  const closeButton = document.querySelector<HTMLButtonElement>('#close-upload-modal')
  const cancelButton = document.querySelector<HTMLButtonElement>('#cancel-upload')
  const submitButton = document.querySelector<HTMLButtonElement>('#submit-upload')
  const input = document.querySelector<HTMLInputElement>('#workbook-upload-input')
  const dropzone = document.querySelector<HTMLElement>('#upload-dropzone')

  if (!modal || !openButton || !closeButton || !cancelButton || !submitButton || !input || !dropzone) return

  openButton.addEventListener('click', () => {
    modal.hidden = false
  })

  closeButton.addEventListener('click', closeUploadModal)
  cancelButton.addEventListener('click', closeUploadModal)
  submitButton.addEventListener('click', () => {
    void uploadSelectedWorkbook()
  })

  modal.addEventListener('click', (event) => {
    if (event.target === modal) closeUploadModal()
  })

  input.addEventListener('change', () => {
    setSelectedUploadFile(input.files?.[0] ?? null)
  })

  dropzone.addEventListener('dragover', (event) => {
    event.preventDefault()
    dropzone.classList.add('is-dragging')
  })

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('is-dragging')
  })

  dropzone.addEventListener('drop', (event) => {
    event.preventDefault()
    dropzone.classList.remove('is-dragging')
    setSelectedUploadFile(event.dataTransfer?.files[0] ?? null)
  })
}

async function loadDashboard() {
  document.body.dataset.view = 'dashboard'

  const dashboard = document.querySelector<HTMLElement>('#dashboard')
  const list = document.querySelector<HTMLElement>('#workbook-list')
  const createButton = document.querySelector<HTMLButtonElement>('#create-workbook')
  const newWorkbookNameInput = document.querySelector<HTMLInputElement>('#new-workbook-name')

  if (!dashboard || !list || !createButton) return

  dashboard.hidden = false
  list.innerHTML = '<div class="workbook-empty">Carregando planilhas...</div>'
  setupUploadModal()

  createButton.addEventListener('click', async () => {
    createButton.disabled = true
    createButton.textContent = 'Criando...'

    try {
      const response = await fetch('/api/workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: normalizeWorkbookName(newWorkbookNameInput?.value || '') }),
      })

      if (!response.ok) throw new Error('Nao foi possivel criar a planilha.')

      const data = (await response.json()) as { workbook: WorkbookListItem }
      window.location.href = workbookUrl(data.workbook.id)
    } catch (error) {
      console.error(error)
      alert('Nao foi possivel criar a planilha agora.')
      createButton.disabled = false
      createButton.textContent = 'Criar nova planilha'
    }
  })

  try {
    const response = await fetch('/api/workbooks')
    if (!response.ok) throw new Error('Nao foi possivel carregar as planilhas.')

    const data = (await response.json()) as { workbooks: WorkbookListItem[] }

    if (data.workbooks.length === 0) {
      list.innerHTML = '<div class="workbook-empty">Nenhuma planilha criada ainda.</div>'
      return
    }

    list.innerHTML = data.workbooks
      .map((workbook) => `
        <div class="workbook-row">
          <a class="workbook-item" href="${workbookUrl(workbook.id)}">
            <strong>${escapeHtml(workbook.name)}</strong>
            <span>Atualizada em ${formatDate(workbook.updatedAt)}</span>
          </a>
          <div class="workbook-row-actions">
            <button
              class="text-button rename-workbook"
              type="button"
              data-workbook-id="${escapeHtml(workbook.id)}"
              data-workbook-name="${escapeHtml(workbook.name)}"
            >
              Renomear
            </button>
            <span>Abrir</span>
          </div>
        </div>
      `)
      .join('')

    list.querySelectorAll<HTMLButtonElement>('.rename-workbook').forEach((button) => {
      button.addEventListener('click', async () => {
        const id = button.dataset.workbookId
        const currentName = button.dataset.workbookName || ''
        const nextName = window.prompt('Novo nome da planilha', currentName)

        if (!id || nextName === null) return

        try {
          await renameWorkbook(id, normalizeWorkbookName(nextName))
          window.location.reload()
        } catch (error) {
          console.error(error)
          alert('Nao foi possivel renomear a planilha agora.')
        }
      })
    })
  } catch (error) {
    console.error(error)
    list.innerHTML = '<div class="workbook-empty">Nao foi possivel carregar as planilhas.</div>'
  }
}

function columnIndexToName(index: number) {
  let value = index + 1
  let name = ''

  while (value > 0) {
    const modulo = (value - 1) % 26
    name = String.fromCharCode(65 + modulo) + name
    value = Math.floor((value - modulo) / 26)
  }

  return name
}

function getStatusColumnRangeNotation(rowCount: number) {
  const columnName = columnIndexToName(STATUS_COLUMN_INDEX)
  return `${columnName}2:${columnName}${Math.max(rowCount, 2)}`
}

function worksheetHasStatusColumn(worksheet: any) {
  const statusHeader = worksheet?.getRange(`${columnIndexToName(STATUS_COLUMN_INDEX)}1`).getValue()

  return String(statusHeader || '').trim() === STATUS_COLUMN_NAME
}

function applyStatusDropdownToWorksheet(univerAPI: ReturnType<typeof setupUniver>, worksheet: any) {
  const newDataValidation = (univerAPI as any).newDataValidation

  if (!worksheetHasStatusColumn(worksheet) || typeof newDataValidation !== 'function') return

  const rowCount = Math.max(worksheet.getMaxRows?.() || 1000, 2)
  const rule = newDataValidation.call(univerAPI)
    .requireValueInList(STATUS_OPTIONS.map((option) => option.label), false, true)
    .setOptions({
      allowBlank: true,
      showErrorMessage: true,
      error: 'Escolha um status da lista.',
      formula2: STATUS_OPTIONS.map((option) => option.color).join(','),
      renderMode: (univerAPI.Enum as any).DataValidationRenderMode?.CUSTOM,
    })
    .build()

  const statusRange = worksheet.getRange(getStatusColumnRangeNotation(rowCount)) as any

  if (typeof statusRange.setDataValidation === 'function') statusRange.setDataValidation(rule)
}

function setupStatusColumn(univerAPI: ReturnType<typeof setupUniver>) {
  const workbook = univerAPI.getActiveWorkbook()
  const sheets = workbook?.getSheets?.() || []

  sheets.forEach((worksheet: any) => {
    applyStatusDropdownToWorksheet(univerAPI, worksheet)
  })
}

async function loadWorkbook(workbookId: string) {
  const response = await fetch(`/api/workbooks/${encodeURIComponent(workbookId)}`)

  if (!response.ok) {
    window.history.replaceState(null, '', '/')
    await loadDashboard()
    alert('Planilha nao encontrada.')
    return
  }

  const data = (await response.json()) as WorkbookResponse

  document.body.dataset.view = 'editor'

  const univerAPI = setupUniver(data.workbook.snapshot)

  // test on dev
  window.univerAPI = univerAPI

  setupQuickToolbar(univerAPI)
  setupAutosave(univerAPI, data.workbook.id)

  univerAPI.addEvent(univerAPI.Event.LifeCycleChanged, ({ stage }) => {
    if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
      document.body.dataset.univerReady = 'true'
      setupStatusColumn(univerAPI)
    }
  })
}

function setupAutosave(univerAPI: ReturnType<typeof setupUniver>, workbookId: string) {
  let lastSaved = ''
  let saving = false

  async function save() {
    const workbook = univerAPI.getActiveWorkbook()
    if (!workbook || saving) return

    const snapshot = workbook.save()
    const serialized = JSON.stringify(snapshot)

    if (serialized === lastSaved) return

    saving = true

    try {
      const response = await fetch(`/api/workbooks/${encodeURIComponent(workbookId)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ snapshot }),
      })

      if (!response.ok) throw new Error('Nao foi possivel salvar a planilha.')
      lastSaved = serialized
    } catch (error) {
      console.error(error)
    } finally {
      saving = false
    }
  }

  window.setInterval(save, 2500)
  window.addEventListener('beforeunload', () => {
    void save()
  })
}

async function main() {
  const params = new URLSearchParams(window.location.search)
  const workbookId = params.get('id')

  if (!workbookId) {
    await loadDashboard()
    return
  }

  await loadWorkbook(workbookId)
}

void main()
