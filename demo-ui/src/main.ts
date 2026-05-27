import './style.css'
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

async function importXlsFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' })
  const snapshot = createDefaultWorkbookData('imported-workbook', workbookNameFromFile(file)) as any

  snapshot.sheetOrder = []
  snapshot.sheets = {}

  workbook.SheetNames.forEach((sheetName, index) => {
    const worksheet = workbook.Sheets[sheetName]
    const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : null
    const sheetId = `sheet-${index + 1}`
    const rowCount = range ? Math.max(range.e.r + 1, 100) : 100
    const columnCount = range ? Math.max(range.e.c + 1, 20) : 20
    const cellData: Record<string, Record<string, Record<string, unknown>>> = {}

    if (range) {
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let column = range.s.c; column <= range.e.c; column += 1) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: column })]
          if (!cell) continue

          cellData[row] ||= {}
          cellData[row][column] = {
            v: cell.w ?? cell.v ?? '',
            ...(cell.f ? { f: cell.f.startsWith('=') ? cell.f : `=${cell.f}` } : {}),
          }
        }
      }
    }

    snapshot.sheetOrder.push(sheetId)
    snapshot.sheets[sheetId] = {
      ...createDefaultWorkbookData(sheetId, sheetName).sheets['sheet-01'],
      id: sheetId,
      name: sheetName || `Sheet${index + 1}`,
      rowCount,
      columnCount,
      cellData,
    }
  })

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
