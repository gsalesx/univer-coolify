import './style.css'
import { setupUniver } from './setup-univer'

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

async function loadDashboard() {
  document.body.dataset.view = 'dashboard'

  const dashboard = document.querySelector<HTMLElement>('#dashboard')
  const list = document.querySelector<HTMLElement>('#workbook-list')
  const createButton = document.querySelector<HTMLButtonElement>('#create-workbook')

  if (!dashboard || !list || !createButton) return

  dashboard.hidden = false
  list.innerHTML = '<div class="workbook-empty">Carregando planilhas...</div>'

  createButton.addEventListener('click', async () => {
    createButton.disabled = true
    createButton.textContent = 'Criando...'

    try {
      const response = await fetch('/api/workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: 'Untitled spreadsheet' }),
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
        <a class="workbook-item" href="${workbookUrl(workbook.id)}">
          <div>
            <strong>${escapeHtml(workbook.name)}</strong>
            <span>Atualizada em ${formatDate(workbook.updatedAt)}</span>
          </div>
          <span>Abrir</span>
        </a>
      `)
      .join('')
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
