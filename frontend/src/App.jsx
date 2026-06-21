import { useState, useEffect, Fragment } from 'react'
import { Container, Row, Col, Card, Button, Form, Alert, Spinner, Table, Badge, Collapse, ProgressBar, Modal } from 'react-bootstrap'
import axios from 'axios'
import './App.css'

function SessionDetails({ data }) {
  if (!data) return <div className="p-2 text-muted">Загрузка...</div>
  const meta = data.meta || {}
  const blocks = meta.blocks || {}
  const renderBlock = (label, files, badgeBg) => (
    <div className="mb-2">
      <strong>{label}:</strong>{' '}
      {files && files.length > 0 ? (
        files.map((f, i) => <Badge bg={badgeBg} key={i} className="me-1 mb-1">{f}</Badge>)
      ) : (
        <span className="text-muted small">нет</span>
      )}
    </div>
  )
  return (
    <div className="p-3">
      <Row>
        <Col md={6}>
          {renderBlock('Чек-лист (блок 1)', blocks.checklist || (meta.template_file ? [meta.template_file] : []), 'primary')}
          {renderBlock('План аудита (блок 2)', blocks.plan || (meta.plan_doc_file ? [meta.plan_doc_file] : []), 'info')}
          {renderBlock('Источники (блок 3)', blocks.sources || meta.source_files || [], 'success')}
        </Col>
        <Col md={6}>
          <div className="mb-2"><strong>Создана:</strong> <span className="small">{meta.created_at || '—'}</span></div>
          <div className="mb-2"><strong>Завершена:</strong> <span className="small">{meta.finished_at || '—'}</span></div>
          <div className="mb-2"><strong>Модель:</strong> <code className="small">{meta.model || '—'}</code></div>
          <div className="mb-2">
            <strong>Все файлы в сессии ({(data.files || []).length}):</strong>
            <div className="small text-muted" style={{ maxHeight: 100, overflowY: 'auto' }}>
              {(data.files || []).map(f => f.name).join(', ') || '—'}
            </div>
          </div>
          {(data.output_files || []).length > 0 && (
            <div className="mb-2">
              <strong>Результат:</strong>{' '}
              {data.output_files.map((f, i) => (
                <a key={i} href={`${API_BASE}/api/download/${encodeURIComponent(f.ref)}`}
                   className="me-2" target="_blank" rel="noreferrer">
                  📥 {f.name}
                </a>
              ))}
            </div>
          )}
        </Col>
      </Row>
      {meta.checklist && meta.checklist.length > 0 && (
        <details className="mt-2">
          <summary className="small text-muted" style={{ cursor: 'pointer' }}>
            Показать вердикты по {meta.checklist.length} пунктам
          </summary>
          <Table size="sm" bordered className="mt-2 mb-0">
            <thead><tr><th style={{ width: '4ch' }}>#</th><th style={{ width: '8ch' }}>Статус</th><th>Обоснование</th></tr></thead>
            <tbody>
              {meta.checklist.map((it, i) => (
                <tr key={i}>
                  <td>{i + 1}</td>
                  <td>{it.ok ? <Badge bg="success">OK</Badge> : <Badge bg="danger">NOK</Badge>}</td>
                  <td className="small">{it.reason || '—'}</td>
                </tr>
              ))}
            </tbody>
          </Table>
        </details>
      )}
    </div>
  )
}

const API_BASE = import.meta.env.VITE_API_BASE ?? ''

// Известные модели GigaChat (фолбэк, если диагностика ещё не запускалась).
// Реальный список приходит с /api/gigachat/diagnose и заменяет этот.
const KNOWN_MODELS = [
  'GigaChat',
  'GigaChat-Lite',
  'GigaChat-Pro',
  'GigaChat-Max',
  'GigaChat-2',
  'GigaChat-2-Pro',
  'GigaChat-2-Max',
]

function App() {
  const [apiKey, setApiKey] = useState('')
  const [apiKeySaved, setApiKeySaved] = useState(false)
  const [saveSuccess, setSaveSuccess] = useState(false)
  const [model, setModel] = useState('')
  const [defaultModel, setDefaultModel] = useState('')
  const [availableModels, setAvailableModels] = useState(KNOWN_MODELS)
  const [diagnosing, setDiagnosing] = useState(false)
  const [diagnose, setDiagnose] = useState(null)
  const [files, setFiles] = useState({ checklists: [], planDocs: [], sources: [] })
  const [uploadedFiles, setUploadedFiles] = useState({ checklists: [], planDocs: [], sources: [] })
  const [uploadProgress, setUploadProgress] = useState({ checklists: null, planDocs: null, sources: null })
  const [processing, setProcessing] = useState(false)
  const [progress, setProgress] = useState(null)
  const [result, setResult] = useState(null)
  const [error, setError] = useState(null)
  const [checklistPreview, setChecklistPreview] = useState(false)
  const [sessionId, setSessionId] = useState(null)
  const [activityLog, setActivityLog] = useState([])
  const [startTime, setStartTime] = useState(null)
  const [elapsed, setElapsed] = useState(0)
  const [showAdmin, setShowAdmin] = useState(false)
  const [showInstructions, setShowInstructions] = useState(false)
  const [adminSessions, setAdminSessions] = useState([])
  const [adminLoading, setAdminLoading] = useState(false)
  const [adminExpanded, setAdminExpanded] = useState({})

  const loadAdminSessions = async () => {
    setAdminLoading(true)
    try {
      const res = await axios.get(`${API_BASE}/api/admin/sessions?limit=100`)
      setAdminSessions(res.data.sessions || [])
    } catch (e) {
      setError('Не удалось загрузить список сессий')
    } finally {
      setAdminLoading(false)
    }
  }

  const toggleAdminDetails = async (sid) => {
    if (adminExpanded[sid]?.loaded) {
      setAdminExpanded(prev => ({ ...prev, [sid]: { ...prev[sid], open: !prev[sid].open } }))
      return
    }
    try {
      const res = await axios.get(`${API_BASE}/api/admin/sessions/${sid}`)
      setAdminExpanded(prev => ({ ...prev, [sid]: { open: true, loaded: true, data: res.data } }))
    } catch (e) {
      setError('Не удалось загрузить детали сессии')
    }
  }

  const downloadSessionZip = (sid) => {
    window.open(`${API_BASE}/api/admin/sessions/${sid}/zip`, '_blank')
  }

  const deleteSession = async (sid) => {
    if (!confirm(`Удалить сессию ${sid.slice(0, 8)}…? Файлы и результат будут стёрты.`)) return
    try {
      await axios.delete(`${API_BASE}/api/admin/sessions/${sid}`)
      setAdminSessions(prev => prev.filter(s => s.session_id !== sid))
    } catch (e) {
      setError('Не удалось удалить сессию')
    }
  }

  useEffect(() => {
    if (showAdmin) loadAdminSessions()
  }, [showAdmin])

  // Загрузка настроек при старте
  useEffect(() => {
    loadSettings()
  }, [])

  // Polling статуса обработки + добавление новых сообщений в лог
  useEffect(() => {
    if (!processing) return
    let lastKey = ''
    const id = setInterval(async () => {
      try {
        const res = await axios.get(`${API_BASE}/api/status`)
        setProgress(res.data)
        const key = `${res.data.stage}|${res.data.message}|${res.data.detail}`
        if (key !== lastKey && res.data.message) {
          lastKey = key
          const ts = new Date().toLocaleTimeString('ru-RU', { hour12: false })
          const entry = `[${ts}] ${res.data.message}${res.data.detail ? ' · ' + res.data.detail : ''}`
          setActivityLog(prev => [...prev.slice(-99), entry])
        }
      } catch (e) {
        // игнорируем — статус необязателен
      }
    }, 700)
    return () => clearInterval(id)
  }, [processing])

  // Таймер: считаем секунды с момента старта обработки
  useEffect(() => {
    if (!processing || !startTime) return
    const id = setInterval(() => {
      setElapsed(Math.floor((Date.now() - startTime) / 1000))
    }, 1000)
    return () => clearInterval(id)
  }, [processing, startTime])

  const formatElapsed = (sec) => {
    const m = Math.floor(sec / 60)
    const s = sec % 60
    return m > 0 ? `${m}м ${s}с` : `${s}с`
  }

  const STAGE_ORDER = [
    { key: 'extract', label: 'Чтение файлов', icon: '📥' },
    { key: 'preflight', label: 'GigaChat', icon: '🔌' },
    { key: 'header', label: 'Шапка', icon: '📋' },
    { key: 'verify', label: 'Проверка пунктов', icon: '🔍' },
    { key: 'fill', label: 'Заполнение', icon: '📝' },
    { key: 'validate', label: 'Валидация', icon: '🧪' },
    { key: 'done', label: 'Готово', icon: '✓' },
  ]
  const stageStatus = (stageKey) => {
    if (!progress) return 'pending'
    const curIdx = STAGE_ORDER.findIndex(s => s.key === progress.stage)
    const idx = STAGE_ORDER.findIndex(s => s.key === stageKey)
    if (progress.stage === 'error') return idx <= curIdx ? 'error' : 'pending'
    if (idx < curIdx) return 'done'
    if (idx === curIdx) return 'active'
    return 'pending'
  }

  const loadSettings = async () => {
    try {
      const res = await axios.get(`${API_BASE}/api/settings/gigachat`)
      if (res.data.api_key) {
        setApiKey(res.data.api_key)
        setApiKeySaved(true)
      }
      if (res.data.default_model) setDefaultModel(res.data.default_model)
      if (res.data.model) {
        setModel(res.data.model)
        setAvailableModels((prev) =>
          prev.includes(res.data.model) ? prev : [...prev, res.data.model]
        )
      }
    } catch (e) {
      console.error('Ошибка загрузки настроек:', e)
    }
  }

  const runDiagnose = async () => {
    setDiagnosing(true)
    setDiagnose(null)
    try {
      const res = await axios.get(`${API_BASE}/api/gigachat/diagnose`, { timeout: 30000 })
      setDiagnose(res.data)
      // Подменяем список моделей на актуальный с ключа
      if (Array.isArray(res.data.models) && res.data.models.length > 0) {
        setAvailableModels(res.data.models)
      }
    } catch (e) {
      setDiagnose({
        ok: false,
        stage: 'frontend_error',
        detail: e.response?.data?.detail || e.message || 'неизвестная ошибка',
        models: [],
      })
    } finally {
      setDiagnosing(false)
    }
  }

  const saveApiKey = async () => {
    if (!apiKey || !apiKey.trim()) {
      setError('Введите API ключ перед сохранением')
      return
    }
    try {
      await axios.post(`${API_BASE}/api/settings/gigachat`, {
        api_key: apiKey.trim(),
        model: (model || '').trim(),
      })
      setApiKeySaved(true)
      setError(null)
      setSaveSuccess(true)
      setTimeout(() => setSaveSuccess(false), 2500)
    } catch (e) {
      const detail = e.response?.data?.detail || e.message || 'сеть недоступна'
      setError(`Ошибка сохранения настроек: ${detail}`)
    }
  }

  const handleFileChange = (type, e) => {
    const selectedFiles = Array.from(e.target.files)
    if (type === 'checklists' || type === 'planDocs') {
      // в этих полях максимум 2 файла
      setFiles(prev => ({ ...prev, [type]: selectedFiles.slice(0, 2) }))
    } else if (type === 'sources') {
      setFiles(prev => ({ ...prev, sources: selectedFiles }))
    }
  }

  const doUpload = async (key, formData, errMsg) => {
    setError(null)
    setUploadProgress(prev => ({ ...prev, [key]: 0 }))
    try {
      const res = await axios.post(`${API_BASE}/api/upload`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (e) => {
          if (e.total) {
            setUploadProgress(prev => ({ ...prev, [key]: Math.round((e.loaded / e.total) * 100) }))
          }
        }
      })
      return res
    } catch (e) {
      setError(errMsg)
      throw e
    } finally {
      setUploadProgress(prev => ({ ...prev, [key]: null }))
    }
  }

  // Универсальная загрузка набора файлов в одно поле.
  // key — ключ в state (checklists | planDocs | sources); max — лимит файлов в поле.
  const uploadMulti = async (key, fileType, block, max) => {
    const sel = files[key] || []
    if (sel.length === 0) {
      setError('Выберите файлы')
      return
    }
    const formData = new FormData()
    sel.forEach(f => formData.append('files', f))
    formData.append('file_type', fileType)
    formData.append('block', block)
    if (sessionId) formData.append('session_id', sessionId)
    try {
      const res = await doUpload(key, formData, 'Ошибка загрузки файлов')
      if (res.data.session_id) setSessionId(res.data.session_id)
      setUploadedFiles(prev => ({
        ...prev,
        [key]: [...prev[key], ...res.data.uploaded_files].slice(0, max)
      }))
      setFiles(prev => ({ ...prev, [key]: [] }))
    } catch (e) {}
  }

  const uploadChecklists = () => uploadMulti('checklists', 'plan', 'checklist', 2)
  const uploadPlanDocs = () => uploadMulti('planDocs', 'plan_doc', 'plan', 2)
  const uploadSources = () => uploadMulti('sources', 'source', 'sources', 999)

  const uploadFromPath = async (type) => {
    const labels = { checklists: 'чек-листами', sources: 'источниками' }
    const max = { checklists: 2, sources: 999 }
    const path = prompt(`Введите путь к папке с ${labels[type]}:\n(например: D:\\Perpektiva\\Пакет 2):`)
    if (!path) return

    try {
      const formData = new FormData()
      formData.append('source_path', path)
      formData.append('file_type', type === 'checklists' ? 'plan' : 'source')
      formData.append('block', type === 'checklists' ? 'checklist' : 'sources')
      if (sessionId) formData.append('session_id', sessionId)
      const res = await axios.post(`${API_BASE}/api/upload-from-path`, formData)
      if (res.data.session_id) setSessionId(res.data.session_id)
      setUploadedFiles(prev => ({
        ...prev,
        [type]: [...prev[type], ...res.data.uploaded_files].slice(0, max[type])
      }))
    } catch (e) {
      setError('Ошибка загрузки из папки. Убедитесь, что путь существует.')
    }
  }

  const removeFile = (type, index) => {
    setUploadedFiles(prev => ({
      ...prev,
      [type]: prev[type].filter((_, i) => i !== index)
    }))
  }

  const processDocuments = async () => {
    if (!apiKey) {
      setError('API ключ GigaChat не сохранён. Откройте «🔧 Админ» и сохраните ключ.')
      return
    }
    if (uploadedFiles.checklists.length === 0) {
      setError('Загрузите чек-лист(ы) в поле 1 («План АУДИТА» и/или «Сводный акт»)')
      return
    }
    if (uploadedFiles.planDocs.length === 0) {
      setError('Загрузите проверяемый документ(ы) в поле 2')
      return
    }
    if (uploadedFiles.sources.length === 0) {
      setError('Загрузите документы-источники данных')
      return
    }
    if (!sessionId) {
      setError('Сессия загрузки не найдена. Загрузите файлы заново.')
      return
    }

    setError(null)
    setResult(null)
    setProgress(null)
    setActivityLog([])
    setStartTime(Date.now())
    setElapsed(0)
    setProcessing(true)

    try {
      const nameOf = (f) => f.filename || f.name
      const cks = uploadedFiles.checklists
      const pds = uploadedFiles.planDocs
      const formData = new FormData()
      formData.append('api_key', apiKey)
      formData.append('session_id', sessionId)
      formData.append('template_file', nameOf(cks[0]))
      if (cks[1]) formData.append('template_file_2', nameOf(cks[1]))
      formData.append('plan_doc_file', nameOf(pds[0]))
      if (pds[1]) formData.append('plan_doc_file_2', nameOf(pds[1]))

      const res = await axios.post(`${API_BASE}/api/process`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 60 * 60 * 1000, // 1 час — per-item пайплайн может идти долго
      })

      setResult(res.data)
    } catch (e) {
      setError(e.response?.data?.detail || 'Ошибка обработки документов')
    } finally {
      setProcessing(false)
    }
  }

  const downloadByRef = async (ref) => {
    if (!ref) return
    try {
      const res = await axios.get(`${API_BASE}/api/download/${ref}`, { responseType: 'blob' })
      const url = window.URL.createObjectURL(new Blob([res.data]))
      const link = document.createElement('a')
      link.href = url
      link.setAttribute('download', ref.split('/').pop() || ref)
      document.body.appendChild(link)
      link.click()
      link.remove()
    } catch (e) {
      setError('Ошибка скачивания файла')
    }
  }

  // Список результатов: новый формат outputs[] либо back-compat одиночный output_file
  const resultOutputs = () => {
    if (result?.outputs && result.outputs.length) return result.outputs
    if (result?.output_file) return [{ type: 'plan', output_file: result.output_file }]
    return []
  }
  const outputLabel = (o) => (o.type === 'svod' ? 'Сводный акт' : 'План АУДИТА')

  return (
    <Container fluid className="py-4">
      <div className="d-flex justify-content-end gap-2 mb-2">
        <Button
          size="sm"
          variant="outline-primary"
          onClick={() => setShowInstructions(true)}
        >
          📖 Инструкция
        </Button>
        <Button
          size="sm"
          variant={showAdmin ? 'dark' : 'outline-dark'}
          onClick={() => setShowAdmin(v => !v)}
        >
          {showAdmin ? '← К рабочему режиму' : '🔧 Админ'}
        </Button>
      </div>

      <Modal show={showInstructions} onHide={() => setShowInstructions(false)} size="lg" scrollable>
        <Modal.Header closeButton>
          <Modal.Title>📖 Инструкция: какие файлы в какое поле</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          <p className="text-muted">
            Приложение заполняет один или два чек-листа аудитора. Тип каждого чек-листа
            («План АУДИТА» или «Сводный акт») определяется автоматически — порядок загрузки не важен.
          </p>

          <Card className="mb-3 border-primary">
            <Card.Header className="bg-primary text-white">📄 Поле 1 — Чек-лист(ы) аудитора</Card.Header>
            <Card.Body>
              <p className="mb-2">Сюда грузим <strong>сам чек-лист</strong> (бланк с пунктами проверки). Можно один или оба:</p>
              <ul className="mb-2">
                <li><strong>«ИИ -ЧК -План АУДИТА»</strong> — чек-лист для Плана аудита;</li>
                <li><strong>«ЧК -Сводный акт»</strong> — чек-лист для Сводного акта.</li>
              </ul>
              <p className="mb-0 small text-muted">
                Нужны оба — выберите сразу два файла в одном поле (до 2 шт.). Нужен один — грузите только его.
                Обработка идёт по очереди: сначала План, затем Сводный акт; на каждый — свой выходной файл.
              </p>
            </Card.Body>
          </Card>

          <Card className="mb-3 border-info">
            <Card.Header className="bg-info text-white">📋 Поле 2 — Проверяемый документ(ы)</Card.Header>
            <Card.Body>
              <p className="mb-2">Главный документ, который проверяем. Зависит от чек-листа:</p>
              <ul className="mb-2">
                <li>для ЧК «План АУДИТА» — файл <strong>«План аудита»</strong> (с разделами «Цель и область», «Основание», «Сроки», «Состав ЭГ»);</li>
                <li>для ЧК «Сводный акт» — файл <strong>«Сводный акт исследования (итог)»</strong>.</li>
              </ul>
              <p className="mb-0 small text-muted">
                Можно выбрать <strong>до 2 документов сразу</strong> в одном поле. Если работаете сразу с двумя
                чек-листами — выберите и «План аудита», и «Сводный акт исследования»; система сама сопоставит
                каждый со своим чек-листом.
              </p>
            </Card.Body>
          </Card>

          <Card className="mb-3 border-success">
            <Card.Header className="bg-success text-white">📚 Поле 3 — Источники данных</Card.Header>
            <Card.Body>
              <p className="mb-1"><strong>Для Плана АУДИТА:</strong></p>
              <ul className="mb-2">
                <li>Договор (включая доп. соглашения), Заявка, Приказ о назначении ЭГ;</li>
                <li>Файлы СТО, Акты предыдущего аудита;</li>
                <li>Орг. структура, Разбивка области по кодам ОКВЭД, Сертификат ИГС, Расчёт трудоёмкости.</li>
              </ul>
              <p className="mb-1"><strong>Для Сводного акта</strong> (каждый файл закрывает свои пункты):</p>
              <ul className="mb-2 small">
                <li><strong>«3.1 Акт Р»</strong> (Акт по результатам аудита, 2 этап) — пункты 3, 11, 14;</li>
                <li><strong>«4.1 Разбивка ОКВЭД»</strong> — пункт 4 (коды ОКВЭД);</li>
                <li><strong>«4.2 Отчёт 1 этапа»</strong> — пункт 8 (численность);</li>
                <li><strong>«6.1 Заявка»</strong> — пункты 6, 7 (площадки, адреса);</li>
                <li><strong>«10.1 Трудоёмкость»</strong> — пункт 10 (режим работы / смены);</li>
                <li><strong>«Шаблон Сводного акта»</strong> — пункт 2 (заполнение всех разделов).</li>
              </ul>
              <p className="mb-0 small text-muted">
                Имена файлов лучше оставлять узнаваемыми (как в примерах выше) — система подбирает нужный источник по имени файла.
                Пункты <strong>9 (Инфраструктура)</strong> и <strong>12 (Несоответствия)</strong> Сводного акта помечаются
                «ПРОВЕРИТЬ ВРУЧНУЮ» — их аудитор сверяет сам.
              </p>
            </Card.Body>
          </Card>

          <Alert variant="secondary" className="mb-0">
            <strong>Порядок работы:</strong> сохраните API-ключ GigaChat в «Админ» → чек-лист(ы) в поле 1 →
            проверяемый документ(ы) в поле 2 → источники в поле 3 → «Запустить обработку».
            На выходе — заполненный файл на каждый чек-лист (OK / NOK / ручные в каждом пункте) и кнопка
            «Скачать оба (ZIP)», если чек-листов два.
          </Alert>
        </Modal.Body>
        <Modal.Footer>
          <Button variant="primary" onClick={() => setShowInstructions(false)}>Понятно</Button>
        </Modal.Footer>
      </Modal>

      <h1 className="mb-4 text-center">
        <Badge bg="primary">План АУДИТА</Badge>
        <br />
        <span className="fs-5">Автоматическое заполнение с GigaChat</span>
      </h1>

      {showAdmin && (
        <Card className="mb-4">
          <Card.Header className="bg-primary text-white">
            🔑 Настройки GigaChat API
          </Card.Header>
          <Card.Body>
            <Row className="g-2 align-items-center">
              <Col md={5}>
                <Form.Label className="small text-muted mb-1">API ключ</Form.Label>
                <Form.Control
                  type="password"
                  value={apiKey}
                  onChange={(e) => { setApiKey(e.target.value); setApiKeySaved(false); setSaveSuccess(false) }}
                  placeholder="Введите ваш API ключ GigaChat"
                />
              </Col>
              <Col md={3}>
                <Form.Label className="small text-muted mb-1">
                  Модель {defaultModel && <span className="text-muted">(по умолч. {defaultModel})</span>}
                </Form.Label>
                <Form.Select
                  value={model || defaultModel || ''}
                  onChange={(e) => { setModel(e.target.value); setApiKeySaved(false); setSaveSuccess(false) }}
                >
                  {availableModels.map((m) => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                </Form.Select>
              </Col>
              <Col md={2}>
                <Form.Label className="small text-muted mb-1">&nbsp;</Form.Label>
                <Button
                  className="w-100"
                  variant={apiKeySaved ? 'outline-primary' : 'primary'}
                  onClick={saveApiKey}
                  disabled={!apiKey || !apiKey.trim()}
                >
                  {apiKeySaved ? '↻ Сохранить' : 'Сохранить'}
                </Button>
              </Col>
              <Col md={2}>
                <Form.Label className="small text-muted mb-1">&nbsp;</Form.Label>
                <Button
                  className="w-100"
                  variant="outline-secondary"
                  onClick={runDiagnose}
                  disabled={diagnosing}
                >
                  {diagnosing ? (
                    <><Spinner animation="border" size="sm" className="me-2" />Проверка...</>
                  ) : (
                    '🔍 Диагностика'
                  )}
                </Button>
              </Col>
            </Row>
            {saveSuccess && (
              <Alert variant="success" className="mt-2 mb-0 py-2">
                ✓ Ключ успешно сохранён на сервере
              </Alert>
            )}
            {apiKeySaved && !saveSuccess && (
              <div className="text-muted small mt-2">
                Ключ уже сохранён. Можно перезаписать при необходимости.
              </div>
            )}
            {diagnose && (
              <Alert
                variant={diagnose.ok ? 'success' : 'warning'}
                className="mt-3 mb-0"
              >
                <div>
                  <strong>Статус: </strong>
                  {diagnose.ok ? '✓ Готов к работе' : `✗ ${diagnose.stage}`}
                </div>
                <div className="mt-1 small">{diagnose.detail}</div>
                {diagnose.current_model && (
                  <div className="mt-1 small">
                    Выбранная модель в backend: <code>{diagnose.current_model}</code>
                  </div>
                )}
                {diagnose.models && diagnose.models.length > 0 && (
                  <div className="mt-1 small">
                    Доступные модели на ключе:{' '}
                    {diagnose.models.map((m) => (
                      <Badge
                        key={m}
                        bg={m === diagnose.current_model ? 'primary' : 'secondary'}
                        className="me-1"
                      >
                        {m}
                      </Badge>
                    ))}
                  </div>
                )}
              </Alert>
            )}
          </Card.Body>
        </Card>
      )}

      {showAdmin && (
        <Card className="mb-4">
          <Card.Header className="bg-dark text-white d-flex justify-content-between align-items-center">
            <span>🗂 История запросов</span>
            <Button size="sm" variant="outline-light" onClick={loadAdminSessions} disabled={adminLoading}>
              {adminLoading ? 'Обновление...' : '↻ Обновить'}
            </Button>
          </Card.Header>
          <Card.Body className="p-0">
            {adminSessions.length === 0 ? (
              <div className="p-3 text-muted">{adminLoading ? 'Загрузка...' : 'Сессий нет.'}</div>
            ) : (
              <Table hover responsive size="sm" className="mb-0">
                <thead className="table-light">
                  <tr>
                    <th style={{ width: '14ch' }}>ID</th>
                    <th>Время</th>
                    <th>Заявитель</th>
                    <th>Файлы</th>
                    <th>Результат</th>
                    <th>Модель</th>
                    <th style={{ width: '20ch' }}>Действия</th>
                  </tr>
                </thead>
                <tbody>
                  {adminSessions.map(s => (
                    <Fragment key={s.session_id}>
                      <tr>
                        <td><code className="small">{s.session_id.slice(0, 8)}…</code></td>
                        <td className="small">{s.finished_at || s.created_at || '—'}</td>
                        <td className="small">{s.applicant || <span className="text-muted">—</span>}</td>
                        <td><Badge bg="secondary">{s.files_count}</Badge></td>
                        <td>
                          {s.total_items != null ? (
                            <>
                              <Badge bg="success" className="me-1">OK {s.ok_count ?? 0}</Badge>
                              <Badge bg="danger">NOK {s.nok_count ?? 0}</Badge>
                            </>
                          ) : (
                            <span className="text-muted small">не обработано</span>
                          )}
                        </td>
                        <td className="small">{s.model || '—'}</td>
                        <td>
                          <Button size="sm" variant="outline-primary" className="me-1" onClick={() => toggleAdminDetails(s.session_id)}>
                            {adminExpanded[s.session_id]?.open ? 'Скрыть' : 'Детали'}
                          </Button>
                          <Button size="sm" variant="outline-success" className="me-1" onClick={() => downloadSessionZip(s.session_id)}>
                            ZIP
                          </Button>
                          <Button size="sm" variant="outline-danger" onClick={() => deleteSession(s.session_id)}>
                            ✕
                          </Button>
                        </td>
                      </tr>
                      {adminExpanded[s.session_id]?.open && (
                        <tr>
                          <td colSpan={7} className="bg-light">
                            <SessionDetails data={adminExpanded[s.session_id]?.data} />
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  ))}
                </tbody>
              </Table>
            )}
          </Card.Body>
        </Card>
      )}

      <Row style={{ display: showAdmin ? 'none' : undefined }}>
        {/* Загрузка файлов - 2 блока: План + Источники */}
        <Col md={4}>
          <Card className="mb-4 h-100">
            <Card.Header className="bg-primary text-white">
              📄 1. Чек-лист аудитора
            </Card.Header>
            <Card.Body className="d-flex flex-column">
              <Form.Label className="text-muted small">
                Бланк(и) с пунктами проверки — «План АУДИТА» и/или «Сводный акт».
                Можно выбрать <strong>до 2 файлов сразу</strong>; тип определяется автоматически.
              </Form.Label>
              <Form.Control
                type="file"
                multiple
                accept=".docx,.doc,.docm,.pdf,.xlsx"
                onChange={(e) => handleFileChange('checklists', e)}
                className="mb-2"
              />
              <div className="d-grid gap-2">
                <Button
                  variant="primary"
                  size="sm"
                  onClick={uploadChecklists}
                  disabled={files.checklists.length === 0 || uploadProgress.checklists !== null}
                >
                  {uploadProgress.checklists !== null ? 'Загрузка...' : `Загрузить чек-лист(ы)${files.checklists.length ? ` (${files.checklists.length})` : ''}`}
                </Button>
                <Button
                  variant="outline-primary"
                  size="sm"
                  onClick={() => uploadFromPath('checklists')}
                  disabled={uploadProgress.checklists !== null}
                >
                  Из папки
                </Button>
              </div>
              {uploadProgress.checklists !== null && (
                <ProgressBar now={uploadProgress.checklists} label={`${uploadProgress.checklists}%`} animated className="mt-2" />
              )}
              {uploadedFiles.checklists.length > 0 && (
                <div className="mt-3">
                  {uploadedFiles.checklists.map((f, idx) => (
                    <Badge key={idx} bg="primary" className="me-2 mb-2 d-inline-flex align-items-center">
                      {f.filename || f.name}
                      <button className="btn-close btn-close-white ms-2" style={{ fontSize: '0.6rem' }} onClick={() => removeFile('checklists', idx)} />
                    </Badge>
                  ))}
                </div>
              )}
            </Card.Body>
          </Card>
        </Col>

        <Col md={4}>
          <Card className="mb-4 h-100">
            <Card.Header className="bg-info text-white">
              📋 2. Проверяемый документ
            </Card.Header>
            <Card.Body className="d-flex flex-column">
              <Form.Label className="text-muted small">
                Главный проверяемый документ: <strong>План аудита</strong> (для ЧК «План АУДИТА») —
                пункты «Цель и область», «Основание», «Сроки», «Состав ЭГ»; либо
                <strong> Сводный акт исследования (итог)</strong> (для ЧК «Сводный акт»).
                Можно выбрать <strong>до 2 файлов сразу</strong>; система сама сопоставит их с чек-листами.
              </Form.Label>
              <Form.Control
                type="file"
                multiple
                accept=".docx,.doc,.docm,.pdf,.xlsx"
                onChange={(e) => handleFileChange('planDocs', e)}
                className="mb-2"
              />
              <div className="d-grid gap-2">
                <Button
                  variant="info"
                  size="sm"
                  onClick={uploadPlanDocs}
                  disabled={files.planDocs.length === 0 || uploadProgress.planDocs !== null}
                >
                  {uploadProgress.planDocs !== null ? 'Загрузка...' : `Загрузить документ(ы)${files.planDocs.length ? ` (${files.planDocs.length})` : ''}`}
                </Button>
              </div>
              {uploadProgress.planDocs !== null && (
                <ProgressBar now={uploadProgress.planDocs} label={`${uploadProgress.planDocs}%`} animated variant="info" className="mt-2" />
              )}
              {uploadedFiles.planDocs.length > 0 && (
                <div className="mt-3">
                  {uploadedFiles.planDocs.map((f, idx) => (
                    <Badge key={idx} bg="info" className="me-2 mb-2 d-inline-flex align-items-center">
                      {f.filename || f.name}
                      <button className="btn-close btn-close-white ms-2" style={{ fontSize: '0.6rem' }} onClick={() => removeFile('planDocs', idx)} />
                    </Badge>
                  ))}
                </div>
              )}
            </Card.Body>
          </Card>
        </Col>

        <Col md={4}>
          <Card className="mb-4 h-100">
            <Card.Header className="bg-success text-white">
              📚 3. Источники данных о компании
            </Card.Header>
            <Card.Body className="d-flex flex-column">
              <Form.Label className="text-muted small">
                Документы для сверки с Планом:
                <ul className="mb-0 ps-3 mt-1">
                  <li>Договор (включая доп. соглашения)</li>
                  <li>Заявка</li>
                  <li>Приказ о назначении ЭГ</li>
                  <li>Файлы СТО</li>
                  <li>Акты предыдущего аудита</li>
                  <li>Орг. структура, Разбивка ОКВЭД, Сертификат ИГС, Расчёт трудоёмкости и пр.</li>
                </ul>
              </Form.Label>
              <Form.Control 
                type="file" 
                multiple 
                accept=".docx,.doc,.docm,.pdf,.xlsx,.xls"
                onChange={(e) => handleFileChange('sources', e)}
                className="mb-2"
              />
              <div className="d-grid gap-2">
                <Button
                  variant="success"
                  size="sm"
                  onClick={uploadSources}
                  disabled={files.sources.length === 0 || uploadProgress.sources !== null}
                >
                  {uploadProgress.sources !== null ? 'Загрузка...' : `Загрузить (${files.sources.length} файлов)`}
                </Button>
                <Button
                  variant="outline-success"
                  size="sm"
                  disabled={uploadProgress.sources !== null}
                  onClick={() => uploadFromPath('sources')}
                >
                  Загрузить папку
                </Button>
              </div>
              {uploadProgress.sources !== null && (
                <ProgressBar now={uploadProgress.sources} label={`${uploadProgress.sources}%`} animated variant="success" className="mt-2" />
              )}
              {uploadedFiles.sources.length > 0 && (
                <div className="mt-3">
                  {uploadedFiles.sources.map((f, idx) => (
                    <Badge key={idx} bg="success" className="me-2 mb-2 d-inline-flex align-items-center">
                      {f.filename || f.name}
                      <button className="btn-close btn-close-white ms-2" style={{ fontSize: '0.6rem' }} onClick={() => removeFile('sources', idx)} />
                    </Badge>
                  ))}
                </div>
              )}
            </Card.Body>
          </Card>
        </Col>

        {/* Основная панель */}
        <Col md={12}>
          {error && <Alert variant="danger">{error}</Alert>}

          {/* Сводка загруженных файлов */}
          <Card className="mb-4">
            <Card.Header className="bg-secondary text-white">
              📋 Сводка
            </Card.Header>
            <Card.Body>
              <Row>
                <Col md={4}>
                  <strong>Чек-лист(ы):</strong>{' '}
                  {uploadedFiles.checklists.length > 0 ? (
                    uploadedFiles.checklists.map((f, i) => (
                      <Badge key={i} bg="primary" className="me-1">{f.filename || f.name}</Badge>
                    ))
                  ) : (
                    <span className="text-muted">не загружен</span>
                  )}
                </Col>
                <Col md={4}>
                  <strong>Проверяемый документ:</strong>{' '}
                  {uploadedFiles.planDocs.length > 0 ? (
                    uploadedFiles.planDocs.map((f, i) => (
                      <Badge key={i} bg="info" className="me-1">{f.filename || f.name}</Badge>
                    ))
                  ) : (
                    <span className="text-muted">не загружен</span>
                  )}
                </Col>
                <Col md={4}>
                  <strong>Источники:</strong>{' '}
                  <Badge bg="success">{uploadedFiles.sources.length} файлов</Badge>
                </Col>
              </Row>
            </Card.Body>
          </Card>

          {/* Кнопка обработки + прогресс */}
          <Card className="mb-4">
            <Card.Body className="text-center">
              <Button
                variant="primary"
                size="lg"
                onClick={processDocuments}
                disabled={processing || uploadedFiles.checklists.length === 0 || uploadedFiles.planDocs.length === 0 || uploadedFiles.sources.length === 0}
              >
                {processing ? (
                  <>
                    <Spinner animation="border" size="sm" className="me-2" />
                    Обработка через {model || defaultModel || 'GigaChat'}...
                  </>
                ) : (
                  '🚀 Запустить обработку и заполнение плана'
                )}
              </Button>

              {processing && (
                <div className="mt-4 text-start">
                  {/* Брэдкрамб стадий */}
                  <div className="d-flex flex-wrap gap-2 mb-3 justify-content-center">
                    {STAGE_ORDER.filter(s => s.key !== 'done').map((s) => {
                      const st = stageStatus(s.key)
                      const bg = st === 'done' ? 'success' : st === 'active' ? 'primary' : st === 'error' ? 'danger' : 'light'
                      const txt = st === 'pending' ? 'text-muted' : ''
                      const icon = st === 'done' ? '✓' : st === 'active' ? '⏳' : s.icon
                      return (
                        <Badge key={s.key} bg={bg} text={st === 'pending' ? 'dark' : undefined}
                               className={`px-3 py-2 ${txt}`} style={{ fontSize: '0.85rem' }}>
                          {icon} {s.label}
                        </Badge>
                      )
                    })}
                  </div>

                  {/* Текущая стадия + таймер */}
                  {progress && (
                    <>
                      <div className="d-flex justify-content-between align-items-center mb-1">
                        <strong style={{ fontSize: '1.05rem' }}>
                          {progress.stage === 'extract' && `📥 Чтение файлов (${progress.current || 0}/${progress.total || 0})`}
                          {progress.stage === 'preflight' && '🔌 Проверка GigaChat'}
                          {progress.stage === 'header' && '📋 Извлечение шапки'}
                          {progress.stage === 'verify' && `🔍 Проверка пунктов (${progress.current}/${progress.total})`}
                          {progress.stage === 'fill' && '📝 Заполнение шаблона'}
                          {progress.stage === 'validate' && '🧪 Валидация результата'}
                          {progress.stage === 'done' && '✓ Завершено'}
                          {progress.stage === 'error' && '✗ Ошибка'}
                          {(!progress.stage || progress.stage === 'idle') && 'Ожидание...'}
                        </strong>
                        <span className="text-muted">
                          ⏱ {formatElapsed(elapsed)}
                          {progress.total > 0 && (progress.stage === 'verify' || progress.stage === 'extract') && (
                            <span className="ms-2">{Math.round((progress.current / progress.total) * 100)}%</span>
                          )}
                        </span>
                      </div>
                      <ProgressBar
                        now={
                          progress.total > 0 && (progress.stage === 'verify' || progress.stage === 'extract')
                            ? (progress.current / progress.total) * 100
                            : (progress.stage === 'done' ? 100 : 0)
                        }
                        animated
                        striped
                        variant={progress.stage === 'error' ? 'danger' : 'primary'}
                      />
                      <div className="mt-2 small">
                        <span className="fw-bold">{progress.message}</span>
                        {progress.detail && <span className="text-muted ms-2">· {progress.detail}</span>}
                      </div>
                    </>
                  )}

                  {/* Живой лог активности */}
                  {activityLog.length > 0 && (
                    <div className="mt-3">
                      <div className="small text-muted mb-1">Лог активности:</div>
                      <div
                        ref={(el) => { if (el) el.scrollTop = el.scrollHeight }}
                        style={{
                          maxHeight: '180px',
                          overflowY: 'auto',
                          background: '#1e1e1e',
                          color: '#d4d4d4',
                          fontFamily: 'Consolas, Monaco, monospace',
                          fontSize: '0.78rem',
                          padding: '0.5rem 0.75rem',
                          borderRadius: '4px',
                          whiteSpace: 'pre-wrap',
                          wordBreak: 'break-word',
                        }}
                      >
                        {activityLog.map((line, i) => (
                          <div key={i}>{line}</div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </Card.Body>
          </Card>

          {/* Результат */}
          {result && (
            <Card className={result.status === 'success' ? 'border-success' : 'border-danger'}>
              <Card.Header className={
                result.status === 'success' ? 'bg-success text-white' : 'bg-danger text-white'
              }>
                {result.status === 'success' ? '✓ Успех' : '✗ Ошибка'}
              </Card.Header>
              <Card.Body>
                <p>{result.message}</p>
                
                {result.analyzed_files && result.analyzed_files.length > 0 && (
                  <Alert variant="info" className="mb-3">
                    <strong>Проанализированные файлы:</strong>
                    <div className="mt-1">
                      {result.analyzed_files.map((f, idx) => (
                        <Badge key={idx} bg="info" className="me-1 mb-1">{f}</Badge>
                      ))}
                    </div>
                  </Alert>
                )}
                
                {result.status === 'success' && resultOutputs().length > 0 && (
                  <>
                    {/* Скачивание результатов */}
                    <div className="d-flex flex-wrap gap-2 align-items-center mb-3">
                      {resultOutputs().map((o, idx) => (
                        <Button key={idx} variant="success" size="lg" onClick={() => downloadByRef(o.output_file)}>
                          📥 Скачать: {outputLabel(o)}
                          {typeof o.ok_count === 'number' && (
                            <span className="ms-2 small">
                              (OK {o.ok_count} / NOK {o.nok_count}{o.manual_count ? ` / ручных ${o.manual_count}` : ''})
                            </span>
                          )}
                        </Button>
                      ))}
                      {resultOutputs().length > 1 && sessionId && (
                        <Button variant="outline-success" size="lg"
                          onClick={() => window.open(`${API_BASE}/api/download-outputs/${sessionId}`, '_blank')}>
                          🗜 Скачать оба (ZIP)
                        </Button>
                      )}
                    </div>

                    {/* Детальный результат по каждому чек-листу */}
                    {resultOutputs().map((o, oi) => {
                      const cl = o.checklist || (oi === 0 ? result.extracted_data?.checklist : null) || []
                      const hdr = o.header || (oi === 0 ? result.extracted_data?.header : null)
                      const v = o.validation || (oi === 0 ? result.validation : null)
                      const notes = v && Array.isArray(v.notes) ? v.notes : []
                      const issues = v && Array.isArray(v.issues) ? v.issues : []
                      return (
                        <Card key={oi} className="mb-3">
                          <Card.Header className="d-flex justify-content-between align-items-center">
                            <span><strong>{outputLabel(o)}</strong>{o.checklist_file ? ` — ${o.checklist_file}` : ''}</span>
                            <span>
                              <Badge bg="success" className="me-1">OK {o.ok_count ?? cl.filter(i => i.ok).length}</Badge>
                              <Badge bg="danger" className="me-1">NOK {o.nok_count ?? cl.filter(i => i.nok).length}</Badge>
                              {(o.manual_count || cl.filter(i => !i.ok && !i.nok).length) > 0 && (
                                <Badge bg="secondary">ручных {o.manual_count ?? cl.filter(i => !i.ok && !i.nok).length}</Badge>
                              )}
                            </span>
                          </Card.Header>
                          <Card.Body>
                            {hdr && Object.keys(hdr).length > 0 && (
                              <Table striped size="sm" className="mb-3">
                                <tbody>
                                  {Object.entries(hdr).map(([key, value]) => (
                                    <tr key={key}>
                                      <td style={{ width: '40%' }}><strong>{key}</strong></td>
                                      <td>{typeof value === 'string' ? value : JSON.stringify(value)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </Table>
                            )}
                            <Table striped size="sm" className="mb-2">
                              <thead>
                                <tr>
                                  <th style={{ width: '5%' }}>#</th>
                                  <th style={{ width: '12%' }}>Статус</th>
                                  <th>Обоснование</th>
                                </tr>
                              </thead>
                              <tbody>
                                {cl.map((item, idx) => (
                                  <tr key={idx}>
                                    <td>{idx + 1}</td>
                                    <td>
                                      {item.nok ? (
                                        <Badge bg="danger">☒ NOK</Badge>
                                      ) : item.ok ? (
                                        <Badge bg="success">☑ OK</Badge>
                                      ) : (
                                        <Badge bg="secondary">✋ Ручная</Badge>
                                      )}
                                    </td>
                                    <td className={item.nok ? 'text-danger' : (!item.ok && !item.nok ? 'text-muted' : '')}>
                                      {item.reason || item.problems || '—'}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </Table>
                            {issues.length > 0 && (
                              <Alert variant="warning" className="mt-2 mb-0">
                                <strong>Проблемы валидации:</strong>
                                <ul className="mb-0">{issues.map((s, i) => <li key={i}>{s}</li>)}</ul>
                              </Alert>
                            )}
                            {notes.length > 0 && (
                              <Alert variant="info" className="mt-2 mb-0">
                                <strong>Примечания:</strong>
                                <ul className="mb-0">{notes.map((s, i) => <li key={i}>{s}</li>)}</ul>
                              </Alert>
                            )}
                          </Card.Body>
                        </Card>
                      )
                    })}
                  </>
                )}
              </Card.Body>
            </Card>
          )}
        </Col>
      </Row>
    </Container>
  )
}

export default App
