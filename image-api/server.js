const crypto = require('node:crypto')
const path = require('node:path')

const { CreateBucketCommand, GetObjectCommand, PutObjectCommand, S3Client } = require('@aws-sdk/client-s3')
const cors = require('cors')
const express = require('express')
const multer = require('multer')

const port = Number(process.env.PORT || 8080)
const bucket = process.env.S3_BUCKET || 'univer-images'
const maxFileSizeMb = Number(process.env.MAX_FILE_SIZE_MB || 10)

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

app.use(cors({ origin: process.env.CORS_ORIGIN || '*' }))

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

app.get('/health', (_req, res) => {
  res.json({ ok: true })
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

ensureBucket().then(() => {
  app.listen(port, () => {
    console.log(`Image API listening on :${port}`)
  })
}).catch((error) => {
  console.error(error)
  process.exit(1)
})
