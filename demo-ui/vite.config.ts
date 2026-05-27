import { randomUUID } from 'node:crypto'
import type { IncomingMessage, ServerResponse } from 'node:http'
import process from 'node:process'
import { defineConfig, loadEnv, type Plugin } from 'vite'

import packageJson from './package.json'

type MockWorkbook = {
  id: string
  name: string
  snapshot: Record<string, unknown>
  createdAt: string
  updatedAt: string
}

function createMockWorkbookSnapshot(id: string, name: string) {
  return {
    id,
    locale: 'enUS',
    name,
    appVersion: '0.24.0',
    sheetOrder: ['sheet-01'],
    sheets: {
      'sheet-01': {
        type: 0,
        id: 'sheet-01',
        name: 'Sheet1',
        cellData: {},
        hidden: 0,
        rowCount: 1000,
        columnCount: 20,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 93,
        defaultRowHeight: 27,
        status: 1,
        showGridlines: 1,
        hideRow: [],
        hideColumn: [],
        rowHeader: {
          width: 46,
          hidden: 0,
        },
        columnHeader: {
          height: 20,
          hidden: 0,
        },
        selections: ['A1'],
        rightToLeft: 0,
        pluginMeta: {},
      },
    },
  }
}

function readJsonBody(req: IncomingMessage) {
  return new Promise<Record<string, unknown>>((resolve, reject) => {
    let body = ''

    req.on('data', (chunk) => {
      body += chunk
    })

    req.on('end', () => {
      if (!body) {
        resolve({})
        return
      }

      try {
        resolve(JSON.parse(body))
      } catch (error) {
        reject(error)
      }
    })

    req.on('error', reject)
  })
}

function sendJson(res: ServerResponse, statusCode: number, data: unknown) {
  res.statusCode = statusCode
  res.setHeader('content-type', 'application/json')
  res.end(JSON.stringify(data))
}

function localWorkbookApiPlugin(): Plugin {
  const workbooks = new Map<string, MockWorkbook>()

  return {
    name: 'local-workbook-api',
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        if (!req.url?.startsWith('/api/workbooks')) {
          next()
          return
        }

        const url = new URL(req.url, 'http://localhost')
        const method = req.method || 'GET'
        const path = url.pathname.replace(/^\/api\/workbooks\/?/, '')

        try {
          if (method === 'GET' && path === '') {
            sendJson(res, 200, {
              workbooks: Array.from(workbooks.values())
                .map(({ id, name, createdAt, updatedAt }) => ({ id, name, createdAt, updatedAt }))
                .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)),
            })
            return
          }

          if (method === 'POST' && path === '') {
            const body = await readJsonBody(req)
            const id = `workbook-${randomUUID()}`
            const name = typeof body.name === 'string' && body.name.trim() ? body.name.trim() : 'Untitled spreadsheet'
            const now = new Date().toISOString()
            const workbook = {
              id,
              name,
              snapshot: createMockWorkbookSnapshot(id, name),
              createdAt: now,
              updatedAt: now,
            }

            workbooks.set(id, workbook)
            sendJson(res, 201, { workbook: { id, name, createdAt: now, updatedAt: now } })
            return
          }

          if (method === 'POST' && path === 'import') {
            const body = await readJsonBody(req)
            const snapshot = body.snapshot

            if (!snapshot || typeof snapshot !== 'object') {
              sendJson(res, 400, { error: 'Envie um snapshot valido.' })
              return
            }

            const id = `workbook-${randomUUID()}`
            const name = typeof body.name === 'string' && body.name.trim()
              ? body.name.trim()
              : typeof (snapshot as Record<string, unknown>).name === 'string'
                  && ((snapshot as Record<string, unknown>).name as string).trim()
                ? ((snapshot as Record<string, unknown>).name as string).trim()
                : 'Untitled spreadsheet'
            const now = new Date().toISOString()
            const workbookSnapshot = { ...(snapshot as Record<string, unknown>), id, name }

            workbooks.set(id, {
              id,
              name,
              snapshot: workbookSnapshot,
              createdAt: now,
              updatedAt: now,
            })
            sendJson(res, 201, { workbook: { id, name, createdAt: now, updatedAt: now } })
            return
          }

          const workbook = workbooks.get(path)

          if (!workbook) {
            sendJson(res, 404, { error: 'Planilha nao encontrada.' })
            return
          }

          if (method === 'GET') {
            sendJson(res, 200, { workbook })
            return
          }

          if (method === 'PATCH') {
            const body = await readJsonBody(req)
            const name = typeof body.name === 'string' && body.name.trim() ? body.name.trim() : ''

            if (!name) {
              sendJson(res, 400, { error: 'Envie um nome valido.' })
              return
            }

            workbook.name = name
            workbook.snapshot = { ...workbook.snapshot, name }
            workbook.updatedAt = new Date().toISOString()
            sendJson(res, 200, {
              workbook: {
                id: workbook.id,
                name: workbook.name,
                createdAt: workbook.createdAt,
                updatedAt: workbook.updatedAt,
              },
            })
            return
          }

          if (method === 'PUT') {
            const body = await readJsonBody(req)
            const snapshot = body.snapshot

            if (!snapshot || typeof snapshot !== 'object') {
              sendJson(res, 400, { error: 'Envie um snapshot valido.' })
              return
            }

            const snapshotRecord = snapshot as Record<string, unknown>

            workbook.snapshot = snapshotRecord
            workbook.name = typeof snapshotRecord.name === 'string' && snapshotRecord.name.trim()
              ? snapshotRecord.name.trim()
              : workbook.name
            workbook.updatedAt = new Date().toISOString()
            sendJson(res, 200, {
              workbook: {
                id: workbook.id,
                name: workbook.name,
                createdAt: workbook.createdAt,
                updatedAt: workbook.updatedAt,
              },
            })
            return
          }

          sendJson(res, 405, { error: 'Metodo nao permitido.' })
        } catch (error) {
          console.error(error)
          sendJson(res, 500, { error: 'Falha no mock local.' })
        }
      })
    },
  }
}

export default ({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return defineConfig({
    plugins: [localWorkbookApiPlugin()],
    server: {
      cors: true,
      proxy: {
        '/universer-api': {
          target: env.UNIVER_ENDPOINT,
          changeOrigin: true,
          ws: true,
        },
      },
      allowedHosts: ['local.univer.plus'],
    },
    define: {
      'process.env.UNIVER_CLIENT_LICENSE': `"${env.UNIVER_CLIENT_LICENSE || '%%UNIVER_CLIENT_LICENSE_PLACEHOLDER%%'}"`,
      'process.env.UNIVER_VERSION': `"${packageJson.dependencies['@univerjs/presets']}"`,
    },
    base: './',
    worker: {
      rollupOptions: {
        output: {
          entryFileNames: 'worker.js',
        },
      },
    },
    build: {
      rollupOptions: {
        output: {
          entryFileNames: 'assets/[name]-[hash].js',
        },
      },
    },
  })
}
