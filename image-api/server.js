const crypto = require('node:crypto')
const path = require('node:path')

const { CreateBucketCommand, GetObjectCommand, PutObjectCommand, S3Client } = require('@aws-sdk/client-s3')
const cors = require('cors')
const express = require('express')
const multer = require('multer')
const { Pool } = require('pg')

const port = Number(process.env.PORT || 8080)
const bucket = process.env.S3_BUCKET || 'univer-images'
const maxFileSizeMb = Number(process.env.MAX_FILE_SIZE_MB || 10)
const databaseUrl = process.env.DATABASE_URL || {
  host: process.env.DATABASE_HOST || 'univer-postgres',
  port: Number(process.env.DATABASE_PORT || 5432),
  user: process.env.DATABASE_USERNAME || 'postgres',
  password: process.env.DATABASE_PASSWORD || 'postgres',
  database: process.env.DATABASE_DBNAME || 'univer',
}

const s3 = new S3Client({
  endpoint: process.env.S3_ENDPOINT || 'http://univer-minio:9000',
  region: process.env.S3_REGION || 'us-east-1',
  forcePathStyle: true,
  credentials: {
    accessKeyId: process.env.S3_ACCESS_KEY_ID || process.env.S3_USER || 'minio',
    secretAccessKey: process.env.S3_SECRET_ACCESS_KEY || process.env.S3_PASSWORD || 'minio123456',
  },
})

const app = express()
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: maxFileSizeMb * 1024 * 1024,
  },
})
const pool = new Pool(typeof databaseUrl === 'string' ? { connectionString: databaseUrl } : databaseUrl)

app.use(cors({ origin: process.env.CORS_ORIGIN || '*' }))
app.use(express.json({ limit: process.env.JSON_BODY_LIMIT || '25mb' }))

function publicBaseUrl(req) {
  if (process.env.PUBLIC_BASE_URL) return process.env.PUBLIC_BASE_URL.replace(/\/$/, '')

  const protocol = req.get('x-forwarded-proto') || req.protocol
  const host = req.get('x-forwarded-host') || req.get('host')
  return `${protocol}://${host}`
}

function publicImageUrl(req, key) {
  return `${publicBaseUrl(req)}/api/images/${encodeURIComponent(key).replace(/%2F/g, '/')}`
}

async function ensureBucket() {
  try {
    await s3.send(new CreateBucketCommand({ Bucket: bucket }))
  } catch (error) {
    const okCodes = ['BucketAlreadyOwnedByYou', 'BucketAlreadyExists']
    if (!okCodes.includes(error.name)) throw error
  }
}

async function ensureDatabase() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS workbooks (
      id text PRIMARY KEY,
      name text NOT NULL,
      snapshot jsonb NOT NULL,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `)
}

function createWorkbookSnapshot(id, name) {
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

app.get('/health', (_req, res) => {
  res.json({ ok: true })
})

app.get('/api/workbooks', async (_req, res, next) => {
  try {
    const { rows } = await pool.query(`
      SELECT id, name, created_at AS "createdAt", updated_at AS "updatedAt"
      FROM workbooks
      ORDER BY updated_at DESC
    `)

    res.json({ workbooks: rows })
  } catch (error) {
    next(error)
  }
})

app.post('/api/workbooks', async (req, res, next) => {
  try {
    const id = `workbook-${crypto.randomUUID()}`
    const name = typeof req.body?.name === 'string' && req.body.name.trim()
      ? req.body.name.trim()
      : 'Untitled spreadsheet'
    const snapshot = createWorkbookSnapshot(id, name)

    const { rows } = await pool.query(`
      INSERT INTO workbooks (id, name, snapshot)
      VALUES ($1, $2, $3)
      RETURNING id, name, created_at AS "createdAt", updated_at AS "updatedAt"
    `, [id, name, snapshot])

    res.status(201).json({ workbook: rows[0] })
  } catch (error) {
    next(error)
  }
})

app.get('/api/workbooks/:id', async (req, res, next) => {
  try {
    const { rows } = await pool.query(`
      SELECT id, name, snapshot, created_at AS "createdAt", updated_at AS "updatedAt"
      FROM workbooks
      WHERE id = $1
    `, [req.params.id])

    if (!rows[0]) {
      res.status(404).json({ error: 'Planilha nao encontrada.' })
      return
    }

    res.json({ workbook: rows[0] })
  } catch (error) {
    next(error)
  }
})

app.put('/api/workbooks/:id', async (req, res, next) => {
  try {
    const snapshot = req.body?.snapshot

    if (!snapshot || typeof snapshot !== 'object') {
      res.status(400).json({ error: 'Envie um snapshot valido.' })
      return
    }

    const name = typeof snapshot.name === 'string' && snapshot.name.trim()
      ? snapshot.name.trim()
      : 'Untitled spreadsheet'

    const { rows } = await pool.query(`
      UPDATE workbooks
      SET name = $2, snapshot = $3, updated_at = now()
      WHERE id = $1
      RETURNING id, name, created_at AS "createdAt", updated_at AS "updatedAt"
    `, [req.params.id, name, snapshot])

    if (!rows[0]) {
      res.status(404).json({ error: 'Planilha nao encontrada.' })
      return
    }

    res.json({ workbook: rows[0] })
  } catch (error) {
    next(error)
  }
})

app.post('/api/images/upload', upload.single('image'), async (req, res, next) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'Envie o arquivo no campo multipart "image".' })
      return
    }

    if (!req.file.mimetype.startsWith('image/')) {
      res.status(400).json({ error: 'O arquivo enviado precisa ser uma imagem.' })
      return
    }

    const extension = path.extname(req.file.originalname || '').toLowerCase() || '.jpg'
    const key = `uploads/${Date.now()}-${crypto.randomUUID()}${extension}`

    await s3.send(new PutObjectCommand({
      Bucket: bucket,
      Key: key,
      Body: req.file.buffer,
      ContentType: req.file.mimetype,
    }))

    res.status(201).json({
      key,
      url: publicImageUrl(req, key),
    })
  } catch (error) {
    next(error)
  }
})

app.get(/^\/api\/images\/(.+)$/, async (req, res, next) => {
  try {
    const key = decodeURIComponent(req.params[0])
    const object = await s3.send(new GetObjectCommand({ Bucket: bucket, Key: key }))

    if (object.ContentType) res.setHeader('content-type', object.ContentType)
    res.setHeader('cache-control', 'public, max-age=31536000, immutable')
    object.Body.pipe(res)
  } catch (error) {
    next(error)
  }
})

app.use((error, _req, res, _next) => {
  console.error(error)
  res.status(500).json({ error: 'Falha ao processar imagem.' })
})

Promise.all([ensureBucket(), ensureDatabase()]).then(() => {
  app.listen(port, () => {
    console.log(`Image API listening on :${port}`)
  })
}).catch((error) => {
  console.error(error)
  process.exit(1)
})
