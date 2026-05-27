import './style.css'
import { setupUniver } from './setup-univer'

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

function main() {
  const univerAPI = setupUniver()

  // test on dev
  window.univerAPI = univerAPI

  setupQuickToolbar(univerAPI)

  univerAPI.addEvent(univerAPI.Event.LifeCycleChanged, ({ stage }) => {
    if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
      document.body.dataset.univerReady = 'true'
    }
  })
}

main()
