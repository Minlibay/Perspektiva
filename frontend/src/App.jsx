import { useState, useEffect } from 'react'
import { Container, Row, Col, Card, Button, Form, Alert, Spinner, Table, Badge, Collapse, ProgressBar } from 'react-bootstrap'
import axios from 'axios'
import './App.css'

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
  const [files, setFiles] = useState({ plan: null, planDoc: null, sources: [] })
  const [uploadedFiles, setUploadedFiles] = useState({ plan: null, planDoc: null, sources: [] })
  const [uploadProgress, setUploadProgress] = useState({ plan: null, planDoc: null, sources: null })
  const [processing, setProcessing] = useState(false)
  const [progress, setProgress] = useState(null)
  const [result, setResult] = useState(null)
  const [error, setError] = useState(null)
  const [checklistPreview, setChecklistPreview] = useState(false)
  const [sessionId, setSessionId] = useState(null)
  const [activityLog, setActivityLog] = useState([])
  const [startTime, setStartTime] = useState(null)
  const [elapsed, setElapsed] = useState(0)

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
    if (type === 'plan') {
      setFiles(prev => ({ ...prev, plan: selectedFiles[0] || null }))
    } else if (type === 'planDoc') {
      setFiles(prev => ({ ...prev, planDoc: selectedFiles[0] || null }))
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

  const uploadPlanDoc = async () => {
    if (!files.planDoc) {
      setError('Выберите файл "План"')
      return
    }
    const formData = new FormData()
    formData.append('files', files.planDoc)
    formData.append('file_type', 'plan_doc')
    if (sessionId) formData.append('session_id', sessionId)
    try {
      const res = await doUpload('planDoc', formData, 'Ошибка загрузки файла "План"')
      if (res.data.session_id) setSessionId(res.data.session_id)
      setUploadedFiles(prev => ({ ...prev, planDoc: res.data.uploaded_files[0] }))
      setFiles(prev => ({ ...prev, planDoc: null }))
    } catch (e) {}
  }

  const uploadPlan = async () => {
    if (!files.plan) {
      setError('Выберите файл Плана АУДИТА')
      return
    }
    const formData = new FormData()
    formData.append('files', files.plan)
    formData.append('file_type', 'plan')
    if (sessionId) formData.append('session_id', sessionId)
    try {
      const res = await doUpload('plan', formData, 'Ошибка загрузки плана')
      if (res.data.session_id) setSessionId(res.data.session_id)
      setUploadedFiles(prev => ({ ...prev, plan: res.data.uploaded_files[0] }))
      setFiles(prev => ({ ...prev, plan: null }))
    } catch (e) {}
  }

  const uploadSources = async () => {
    if (files.sources.length === 0) {
      setError('Выберите файлы источников')
      return
    }
    const formData = new FormData()
    files.sources.forEach(file => formData.append('files', file))
    formData.append('file_type', 'source')
    if (sessionId) formData.append('session_id', sessionId)
    try {
      const res = await doUpload('sources', formData, 'Ошибка загрузки источников')
      if (res.data.session_id) setSessionId(res.data.session_id)
      setUploadedFiles(prev => ({
        ...prev,
        sources: [...prev.sources, ...res.data.uploaded_files]
      }))
      setFiles(prev => ({ ...prev, sources: [] }))
    } catch (e) {}
  }

  const uploadFromPath = async (type) => {
    const labels = { plan: 'Планом АУДИТА', sources: 'источниками' }
    const path = prompt(`Введите путь к папке с ${labels[type]}:\n(например: D:\\Perpektiva\\Пакет 2):`)
    if (!path) return

    try {
      const formData = new FormData()
      formData.append('source_path', path)
      formData.append('file_type', type)
      if (sessionId) formData.append('session_id', sessionId)
      const res = await axios.post(`${API_BASE}/api/upload-from-path`, formData)
      if (res.data.session_id) setSessionId(res.data.session_id)

      if (type === 'sources') {
        setUploadedFiles(prev => ({ 
          ...prev, 
          sources: [...prev.sources, ...res.data.uploaded_files] 
        }))
      } else {
        setUploadedFiles(prev => ({ ...prev, plan: res.data.uploaded_files[0] }))
      }
    } catch (e) {
      setError('Ошибка загрузки из папки. Убедитесь, что путь существует.')
    }
  }

  const removeFile = (type, index = null) => {
    if (type === 'plan') {
      setUploadedFiles(prev => ({ ...prev, plan: null }))
    } else if (type === 'planDoc') {
      setUploadedFiles(prev => ({ ...prev, planDoc: null }))
    } else if (type === 'sources' && index !== null) {
      setUploadedFiles(prev => ({
        ...prev,
        sources: prev.sources.filter((_, i) => i !== index)
      }))
    }
  }

  const processDocuments = async () => {
    if (!apiKey) {
      setError('Введите API ключ GigaChat')
      return
    }
    if (!uploadedFiles.plan) {
      setError('Загрузите файл "ИИ -ЧК -План АУДИТА"')
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
      const formData = new FormData()
      formData.append('api_key', apiKey)
      formData.append('template_file', uploadedFiles.plan.filename || uploadedFiles.plan.name)
      formData.append('session_id', sessionId)
      if (uploadedFiles.planDoc) {
        formData.append('plan_doc_file', uploadedFiles.planDoc.filename || uploadedFiles.planDoc.name)
      }

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

  const downloadResult = async () => {
    if (!result?.output_file) return

    try {
      const res = await axios.get(`${API_BASE}/api/download/${result.output_file}`, {
        responseType: 'blob'
      })
      
      const url = window.URL.createObjectURL(new Blob([res.data]))
      const link = document.createElement('a')
      link.href = url
      link.setAttribute('download', result.output_file)
      document.body.appendChild(link)
      link.click()
      link.remove()
    } catch (e) {
      setError('Ошибка скачивания файла')
    }
  }

  return (
    <Container fluid className="py-4">
      <h1 className="mb-4 text-center">
        <Badge bg="primary">План АУДИТА</Badge>
        <br />
        <span className="fs-5">Автоматическое заполнение с GigaChat</span>
      </h1>

      <Row>
        {/* Настройки GigaChat */}
        <Col md={12} className="mb-4">
          <Card>
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
        </Col>

        {/* Загрузка файлов - 2 блока: План + Источники */}
        <Col md={4}>
          <Card className="mb-4 h-100">
            <Card.Header className="bg-primary text-white">
              📄 1. План АУДИТА (чек-лист)
            </Card.Header>
            <Card.Body className="d-flex flex-column">
              <Form.Label className="text-muted small">Шаблон с пунктами проверки</Form.Label>
              <Form.Control
                type="file"
                accept=".docx,.doc,.docm,.pdf,.xlsx"
                onChange={(e) => handleFileChange('plan', e)}
                className="mb-2"
              />
              <div className="d-grid gap-2">
                <Button
                  variant="primary"
                  size="sm"
                  onClick={uploadPlan}
                  disabled={!files.plan || uploadProgress.plan !== null}
                >
                  {uploadProgress.plan !== null ? 'Загрузка...' : 'Загрузить чек-лист'}
                </Button>
                <Button
                  variant="outline-primary"
                  size="sm"
                  onClick={() => uploadFromPath('plan')}
                  disabled={uploadProgress.plan !== null}
                >
                  Из папки
                </Button>
              </div>
              {uploadProgress.plan !== null && (
                <ProgressBar now={uploadProgress.plan} label={`${uploadProgress.plan}%`} animated className="mt-2" />
              )}
              {uploadedFiles.plan && (
                <div className="mt-3">
                  <Badge bg="primary" className="me-2 mb-2 d-inline-flex align-items-center">
                    {uploadedFiles.plan.filename || uploadedFiles.plan.name}
                    <button className="btn-close btn-close-white ms-2" style={{ fontSize: '0.6rem' }} onClick={() => removeFile('plan')} />
                  </Badge>
                </div>
              )}
            </Card.Body>
          </Card>
        </Col>

        <Col md={4}>
          <Card className="mb-4 h-100">
            <Card.Header className="bg-info text-white">
              📋 2. Документы по аудиту
            </Card.Header>
            <Card.Body className="d-flex flex-column">
              <Form.Label className="text-muted small">Даты аудита, описание, состав ЭГ</Form.Label>
              <Form.Control
                type="file"
                accept=".docx,.doc,.docm,.pdf,.xlsx"
                onChange={(e) => handleFileChange('planDoc', e)}
                className="mb-2"
              />
              <div className="d-grid gap-2">
                <Button
                  variant="info"
                  size="sm"
                  onClick={uploadPlanDoc}
                  disabled={!files.planDoc || uploadProgress.planDoc !== null}
                >
                  {uploadProgress.planDoc !== null ? 'Загрузка...' : 'Загрузить План'}
                </Button>
              </div>
              {uploadProgress.planDoc !== null && (
                <ProgressBar now={uploadProgress.planDoc} label={`${uploadProgress.planDoc}%`} animated variant="info" className="mt-2" />
              )}
              {uploadedFiles.planDoc && (
                <div className="mt-3">
                  <Badge bg="info" className="me-2 mb-2 d-inline-flex align-items-center">
                    {uploadedFiles.planDoc.filename || uploadedFiles.planDoc.name}
                    <button className="btn-close btn-close-white ms-2" style={{ fontSize: '0.6rem' }} onClick={() => removeFile('planDoc')} />
                  </Badge>
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
              <Form.Label className="text-muted small">Все документы с исходными данными компании, по которой проводится аудит</Form.Label>
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
                  Из папки (Пакет 2, 3...)
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
                  <strong>Чек-лист:</strong>{' '}
                  {uploadedFiles.plan ? (
                    <Badge bg="primary">{uploadedFiles.plan.filename || uploadedFiles.plan.name}</Badge>
                  ) : (
                    <span className="text-muted">не загружен</span>
                  )}
                </Col>
                <Col md={4}>
                  <strong>План (документ):</strong>{' '}
                  {uploadedFiles.planDoc ? (
                    <Badge bg="info">{uploadedFiles.planDoc.filename || uploadedFiles.planDoc.name}</Badge>
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
                disabled={processing || !uploadedFiles.plan || uploadedFiles.sources.length === 0}
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
                
                {result.status === 'success' && result.extracted_data && (
                  <>
                    <h5>Извлечённые данные (шапка):</h5>
                    <Table striped size="sm">
                      <tbody>
                        {result.extracted_data.header && Object.entries(result.extracted_data.header).map(([key, value]) => (
                          <tr key={key}>
                            <td style={{ width: '40%' }}><strong>{key}</strong></td>
                            <td>{typeof value === 'string' ? value : JSON.stringify(value)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </Table>
                    
                    <Button 
                      variant="outline-primary" 
                      className="mb-3"
                      onClick={() => setChecklistPreview(!checklistPreview)}
                    >
                      {checklistPreview ? '▼ Скрыть' : '▶ Показать'} чек-лист ({result.extracted_data.checklist?.length || 0} пунктов)
                    </Button>
                    
                    <Collapse in={checklistPreview}>
                      <div>
                        <h5>Результат проверки чек-листа:</h5>
                        <Table striped size="sm" className="mb-3">
                          <thead>
                            <tr>
                              <th style={{ width: '5%' }}>#</th>
                              <th style={{ width: '10%' }}>Статус</th>
                              <th style={{ width: '15%' }}>ИИ данные</th>
                              <th>Обоснование</th>
                            </tr>
                          </thead>
                          <tbody>
                            {result.extracted_data.checklist?.map((item, idx) => (
                              <tr key={idx}>
                                <td>{idx + 1}</td>
                                <td>
                                  {item.nok ? (
                                    <Badge bg="danger">☒ NOK</Badge>
                                  ) : item.ok ? (
                                    <Badge bg="success">☑ OK</Badge>
                                  ) : (
                                    <Badge bg="secondary">?</Badge>
                                  )}
                                </td>
                                <td>
                                  {item.ii_data_found ? (
                                    <Badge bg="info">{item.ii_data_found}</Badge>
                                  ) : (
                                    <span className="text-muted">—</span>
                                  )}
                                </td>
                                <td className={item.nok ? 'text-danger' : ''}>
                                  {item.reason || item.problems || "—"}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </Table>
                        
                        <Row className="mb-3">
                          <Col>
                            <Badge bg="success">
                              OK: {result.extracted_data.checklist?.filter(i => i.ok).length || 0}
                            </Badge>{' '}
                            <Badge bg="danger">
                              NOK: {result.extracted_data.checklist?.filter(i => i.nok).length || 0}
                            </Badge>
                          </Col>
                        </Row>
                      </div>
                    </Collapse>
                    
                    <Button variant="success" size="lg" onClick={downloadResult}>
                      📥 Скачать заполненный план
                    </Button>
                    
                    {result.validation && (
                      <Card className={`mt-3 ${result.validation.valid ? 'border-success' : 'border-danger'}`}>
                        <Card.Header className={result.validation.valid ? 'bg-success text-white' : 'bg-warning text-dark'}>
                          ✓ Результат валидации
                        </Card.Header>
                        <Card.Body>
                          <Row>
                            <Col md={4}>
                              <strong>Шапка:</strong> {result.validation.header_filled || '—'}
                            </Col>
                            <Col md={4}>
                              <strong>Чек-лист:</strong> {result.validation.checklist_total || 0} строк
                            </Col>
                            <Col md={4}>
                              <Badge bg="success" className="me-2">OK: {result.validation.ok_count || 0}</Badge>
                              <Badge bg="danger">NOK: {result.validation.nok_count || 0}</Badge>
                            </Col>
                          </Row>
                          {result.validation.valid ? (
                            <Alert variant="success" className="mt-2 mb-0">
                              ✓ Документ заполнен корректно, проблем не обнаружено
                            </Alert>
                          ) : (
                            <Alert variant="warning" className="mt-2 mb-0">
                              <strong>Обнаружены проблемы:</strong>
                              <ul className="mb-0">
                                {result.validation.issues.map((issue, idx) => (
                                  <li key={idx}>{issue}</li>
                                ))}
                              </ul>
                            </Alert>
                          )}
                        </Card.Body>
                      </Card>
                    )}
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
