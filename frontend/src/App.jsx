import { useState, useEffect, Fragment } from 'react'
import { Container, Row, Col, Card, Button, Form, Alert, Spinner, Table, Badge, ProgressBar, Modal } from 'react-bootstrap'
import axios from 'axios'
import './App.css'

function SessionDetails({ data }) {
  if (!data) return <div className="p-2 text-muted">Загрузка...</div>
  const meta = data.meta || {}
  const blocks = meta.blocks || {}
  const outputs = meta.outputs || []
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
          {renderBlock('Бланки чек-листов (поле 1)', blocks.checklist || (meta.template_file ? [meta.template_file] : []), 'primary')}
          {renderBlock('Проверяемые документы (поле 2)', blocks.plan || (meta.plan_doc_file ? [meta.plan_doc_file] : []), 'info')}
          {renderBlock('Источники (поле 3)', blocks.sources || meta.source_files || [], 'success')}
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
      {outputs.map((o, oi) => (
        <details className="mt-2" key={oi}>
          <summary className="small" style={{ cursor: 'pointer' }}>
            <strong>{o.title || o.type}</strong>{' '}
            {o.error ? <Badge bg="danger">ошибка</Badge> : (
              <>
                <Badge bg="success" className="me-1">OK {o.ok_count}</Badge>
                <Badge bg="danger" className="me-1">NOK {o.nok_count}</Badge>
                {o.manual_count > 0 && <Badge bg="warning" text="dark">вручную {o.manual_count}</Badge>}
              </>
            )}
          </summary>
          {o.error ? <div className="small text-danger mt-1">{o.error}</div> : (
            <Table size="sm" bordered className="mt-2 mb-0">
              <thead><tr><th style={{ width: '4ch' }}>#</th><th style={{ width: '9ch' }}>Статус</th><th>Обоснование</th></tr></thead>
              <tbody>
                {(o.checklist || []).map((it, i) => (
                  <tr key={i}>
                    <td>{i + 1}</td>
                    <td>{it.ok ? <Badge bg="success">OK</Badge> : it.nok ? <Badge bg="danger">NOK</Badge>
                      : <Badge bg="warning" text="dark">вручную</Badge>}</td>
                    <td className="small">{it.reason || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </Table>
          )}
        </details>
      ))}
    </div>
  )
}

const API_BASE = import.meta.env.VITE_API_BASE ?? ''
const SESSION_KEY = 'auditSessionId'

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

// OAuth scope зависит от типа аккаунта, на который выпущен ключ.
const KNOWN_SCOPES = [
  { value: 'GIGACHAT_API_PERS', label: 'PERS — физлицо' },
  { value: 'GIGACHAT_API_B2B', label: 'B2B — организация/ИП (предоплата)' },
  { value: 'GIGACHAT_API_CORP', label: 'CORP — организация/ИП (по факту)' },
]

function App() {
  const [apiKey, setApiKey] = useState('')
  const [apiKeySaved, setApiKeySaved] = useState(false)
  const [saveSuccess, setSaveSuccess] = useState(false)
  const [model, setModel] = useState('')
  const [defaultModel, setDefaultModel] = useState('')
  const [scope, setScope] = useState('')
  const [availableModels, setAvailableModels] = useState(KNOWN_MODELS)
  const [diagnosing, setDiagnosing] = useState(false)
  const [diagnose, setDiagnose] = useState(null)
  const [error, setError] = useState(null)
  const [sessionId, setSessionId] = useState(null)
  const [analysis, setAnalysis] = useState(null)          // разбор пакета: файлы, чек-листы, предупреждения
  const [analyzing, setAnalyzing] = useState(false)
  const [analysisProgress, setAnalysisProgress] = useState(null)
  const [overrides, setOverrides] = useState({})          // {файл: тип документа} — исправления пользователя
  const [choices, setChoices] = useState({})              // {бланк: [проверяемые файлы]}
  const [selected, setSelected] = useState({})            // {бланк: проверять ли}
  const [uploading, setUploading] = useState({})
  const [job, setJob] = useState(null)                    // статус фоновой проверки и результаты
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

  const formatElapsed = (sec) => {
    const m = Math.floor(sec / 60)
    const s = sec % 60
    return m > 0 ? `${m}м ${s}с` : `${s}с`
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
      if (res.data.scope) setScope(res.data.scope)
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
        scope: (scope || '').trim(),
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

  // ===== Пакет аудита: три поля → разбор → выбор чек-листов → фоновая проверка =====
  const FIELDS = [
    { key: 'checklist', title: '📄 1. Бланки чек-листов', variant: 'primary',
      hint: 'Бланки «Чек лист …», «ИИ -План АУДИТА», «ЧК -Сводный акт» — сколько угодно; тип каждого определяется автоматически.' },
    { key: 'plan', title: '📋 2. Проверяемые документы', variant: 'info',
      hint: 'План аудита, Акт, Заключение, ПКД, Протоколы НС, Листы оценки… Система сама сопоставит документ с его чек-листом.' },
    { key: 'sources', title: '📚 3. Источники', variant: 'success',
      hint: 'Договор, Приказ ЭГ, Разбивка ОКВЭД, Акт (для ПКД, ПН, Заключения…), Отчет 1 этапа и пр.' },
  ]

  const ensureSession = async () => {
    if (sessionId) return sessionId
    const res = await axios.post(`${API_BASE}/api/session/new`)
    const sid = res.data.session_id
    setSessionId(sid)
    try { localStorage.setItem(SESSION_KEY, sid) } catch (e) { /* приватный режим */ }
    return sid
  }

  // Разбор пакета. overrides/choices передаём явно — state обновляется асинхронно.
  const runAnalyze = async (sid = sessionId, ov = overrides, ch = choices) => {
    if (!sid) return
    setAnalyzing(true)
    try {
      const res = await axios.post(`${API_BASE}/api/sessions/${sid}/analyze`,
        { overrides: ov, choices: ch }, { timeout: 30 * 60 * 1000 })
      setAnalysis(res.data)
      setSelected(prev => {
        const next = {}
        // по умолчанию отмечено всё готовое; План и Сводный (GigaChat) — только если ключ сохранён
        for (const c of res.data.checklists) {
          next[c.file] = c.file in prev ? prev[c.file] && c.ready : c.ready && (!c.needs_gigachat || apiKeySaved)
        }
        return next
      })
    } catch (e) {
      setError(e.response?.data?.detail || 'Не удалось разобрать пакет')
    } finally {
      setAnalyzing(false)
      setAnalysisProgress(null)
    }
  }

  // Прогресс чтения файлов, пока идёт разбор (OCR сканов может занять минуты)
  useEffect(() => {
    if (!analyzing || !sessionId) return
    const id = setInterval(async () => {
      try {
        const res = await axios.get(`${API_BASE}/api/sessions/${sessionId}/status`)
        if (res.data.analysis?.total) setAnalysisProgress(res.data.analysis)
      } catch (e) { /* статус необязателен */ }
    }, 800)
    return () => clearInterval(id)
  }, [analyzing, sessionId])

  // Ход проверки: опрашиваем статус, пока идёт фоновая проверка
  useEffect(() => {
    if (!job?.running || !sessionId) return
    const id = setInterval(async () => {
      try {
        const res = await axios.get(`${API_BASE}/api/sessions/${sessionId}/status`)
        setJob(res.data)
      } catch (e) { /* повторим */ }
    }, 1000)
    return () => clearInterval(id)
  }, [job?.running, sessionId])

  useEffect(() => {
    if (!job?.running || !job.started_at) return
    const started = new Date(job.started_at).getTime()
    const id = setInterval(() => setElapsed(Math.max(0, Math.floor((Date.now() - started) / 1000))), 1000)
    return () => clearInterval(id)
  }, [job?.running, job?.started_at])

  // Восстановление сессии после перезагрузки страницы
  useEffect(() => {
    let sid = null
    try { sid = localStorage.getItem(SESSION_KEY) } catch (e) { /* нет доступа */ }
    if (!sid) return
    axios.get(`${API_BASE}/api/sessions/${sid}/status`)
      .then(res => {
        setSessionId(sid)
        setJob(res.data)
        runAnalyze(sid, {}, {})
      })
      .catch(() => { try { localStorage.removeItem(SESSION_KEY) } catch (e) { /* ок */ } })
  }, [])

  const uploadToField = async (field, fileList) => {
    const list = Array.from(fileList || [])
    if (!list.length) return
    setError(null)
    setUploading(prev => ({ ...prev, [field]: true }))
    try {
      const sid = await ensureSession()
      const formData = new FormData()
      list.forEach(f => formData.append('files', f))
      formData.append('session_id', sid)
      formData.append('block', field)
      await axios.post(`${API_BASE}/api/upload`, formData, { headers: { 'Content-Type': 'multipart/form-data' } })
      await runAnalyze(sid)
    } catch (e) {
      setError(e.response?.data?.detail || 'Ошибка загрузки файлов')
    } finally {
      setUploading(prev => ({ ...prev, [field]: false }))
    }
  }

  const uploadFolderToField = async (field) => {
    const path = prompt('Путь к папке на сервере (например: D:\\Perspektiva\\Пакет 2):')
    if (!path) return
    try {
      const sid = await ensureSession()
      const formData = new FormData()
      formData.append('source_path', path)
      formData.append('session_id', sid)
      formData.append('block', field)
      await axios.post(`${API_BASE}/api/upload-from-path`, formData)
      await runAnalyze(sid)
    } catch (e) {
      setError(e.response?.status === 404 && e.response?.data?.detail === 'Endpoint disabled'
        ? 'Загрузка из папки на сервере отключена — выберите файлы.'
        : (e.response?.data?.detail || 'Ошибка загрузки из папки'))
    }
  }

  const removeFile = async (name) => {
    try {
      await axios.delete(`${API_BASE}/api/files/${encodeURIComponent(name)}`, { params: { session_id: sessionId } })
      const { [name]: _, ...ov } = overrides
      setOverrides(ov)
      await runAnalyze(sessionId, ov, choices)
    } catch (e) {
      setError(e.response?.data?.detail || 'Не удалось удалить файл')
    }
  }

  const changeType = (name, type) => {
    const ov = { ...overrides, [name]: type }
    setOverrides(ov)
    runAnalyze(sessionId, ov, choices)
  }

  const changeChoice = (ckFile, files) => {
    const ch = { ...choices, [ckFile]: files }
    setChoices(ch)
    runAnalyze(sessionId, overrides, ch)
  }

  const newPackage = () => {
    if (job?.running) return
    try { localStorage.removeItem(SESSION_KEY) } catch (e) { /* ок */ }
    setSessionId(null)
    setAnalysis(null)
    setOverrides({})
    setChoices({})
    setSelected({})
    setJob(null)
    setError(null)
  }

  const checklistsToRun = (analysis?.checklists || []).filter(c => c.ready && selected[c.file])
  const needsKey = checklistsToRun.some(c => c.needs_gigachat) && !apiKeySaved

  const startRun = async () => {
    setError(null)
    try {
      await axios.post(`${API_BASE}/api/sessions/${sessionId}/run`, {
        checklists: checklistsToRun.map(c => c.file), overrides, choices,
      })
      setElapsed(0)
      setJob({ running: true, stage: 'start', message: 'Запуск проверки…', outputs: [],
               started_at: new Date().toISOString() })
    } catch (e) {
      setError(e.response?.data?.detail || 'Не удалось запустить проверку')
    }
  }

  const filesOf = (field) => (analysis?.files || []).filter(f => f.field === field)
  const outputs = job?.outputs || []

  const detectedBadge = (f) => {
    if (f.detected_by === 'override') return <Badge bg="primary" className="ms-1">выбрано вручную</Badge>
    if (f.detected_by === 'content') return <Badge bg="success" className="ms-1">по содержимому</Badge>
    if (f.detected_by === 'name') return <Badge bg="warning" text="dark" className="ms-1">по имени — проверьте</Badge>
    return <Badge bg="danger" className="ms-1">не распознан</Badge>
  }

  const renderFile = (f) => (
    <div key={f.name} className="border-top py-2 small">
      <div className="d-flex align-items-start gap-2">
        <span className="flex-grow-1 text-break">{f.name}</span>
        <button className="btn-close" style={{ fontSize: '0.6rem' }} aria-label="Удалить файл"
                disabled={job?.running} onClick={() => removeFile(f.name)} />
      </div>
      {f.kind === 'checklist' && (
        <div className="mt-1"><Badge bg="primary">ЧК: {f.checklist_title}</Badge>
          <span className="text-muted ms-2">{f.items} пунктов</span></div>
      )}
      {f.kind === 'unknown' && <div className="mt-1"><Badge bg="danger">не распознан как бланк</Badge></div>}
      {f.kind === 'misplaced_checklist' && (
        <div className="mt-1"><Badge bg="warning" text="dark">бланк «{f.checklist_title}» — перенесите в поле 1</Badge></div>
      )}
      {f.kind === 'document' && (
        <>
          <div className="d-flex align-items-center mt-1">
            <Form.Select size="sm" value={f.doc_type || 'unknown'} disabled={job?.running}
                         onChange={(e) => changeType(f.name, e.target.value)} style={{ maxWidth: 260 }}>
              <option value="unknown">— тип не определён —</option>
              {(analysis?.doc_types || []).map(t => <option key={t.key} value={t.key}>{t.title}</option>)}
            </Form.Select>
            {detectedBadge(f)}
          </div>
          {(f.org || f.used_by.length > 0) && (
            <div className="text-muted mt-1">
              {f.org}{f.org && f.used_by.length > 0 && ' · '}
              {f.used_by.length > 0 && `→ ЧК ${f.used_by.map(u => analysis.checklists.find(c => c.file === u)?.title || u).join(', ')}`}
            </div>
          )}
        </>
      )}
    </div>
  )

  const statusBadges = (o) => (
    <>
      <Badge bg="success" className="me-1">OK {o.ok_count}</Badge>
      <Badge bg="danger" className="me-1">NOK {o.nok_count}</Badge>
      {o.manual_count > 0 && <Badge bg="warning" text="dark">вручную {o.manual_count}</Badge>}
    </>
  )

  const renderOutput = (o, oi) => {
    const notes = o.validation?.notes || []
    const issues = o.validation?.issues || []
    return (
      <Card key={oi} className={`mb-3 ${o.error ? 'border-danger' : ''}`}>
        <Card.Header className="d-flex flex-wrap justify-content-between align-items-center gap-2">
          <span>
            <strong>{o.title}</strong>
            <span className="text-muted small ms-2">{(o.checked_files || []).join(', ')}</span>
          </span>
          {!o.error && (
            <span className="d-flex align-items-center gap-2">
              {statusBadges(o)}
              <Button size="sm" variant="success" onClick={() => downloadByRef(o.output_file)}>📥 Скачать</Button>
            </span>
          )}
        </Card.Header>
        <Card.Body className="py-2">
          {o.error && <Alert variant="danger" className="mb-0">{o.error}</Alert>}
          {(o.warnings || []).map((w, i) => <div key={i} className="small text-warning-emphasis">⚠ {w}</div>)}
          {!o.error && (
            <details className="mt-1">
              <summary className="small text-muted" style={{ cursor: 'pointer' }}>Пункты ({o.total_items})</summary>
              {o.header && Object.keys(o.header).length > 0 && (
                <Table striped size="sm" className="mt-2 mb-2">
                  <tbody>
                    {Object.entries(o.header).map(([key, value]) => (
                      <tr key={key}><td style={{ width: '35%' }}><strong>{key}</strong></td>
                        <td>{typeof value === 'string' ? value : JSON.stringify(value)}</td></tr>
                    ))}
                  </tbody>
                </Table>
              )}
              <Table striped size="sm" className="mt-2 mb-0">
                <thead><tr><th style={{ width: '4%' }}>#</th><th style={{ width: '11%' }}>Статус</th><th>Обоснование</th></tr></thead>
                <tbody>
                  {(o.checklist || []).map((item, idx) => (
                    <tr key={idx} className={!item.ok && !item.nok ? 'table-warning' : ''}>
                      <td>{idx + 1}</td>
                      <td>{item.nok ? <Badge bg="danger">NOK</Badge> : item.ok ? <Badge bg="success">OK</Badge>
                        : <Badge bg="warning" text="dark">вручную</Badge>}</td>
                      <td className={item.nok ? 'text-danger small' : 'small'}>{item.reason || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </Table>
            </details>
          )}
          {issues.length > 0 && (
            <Alert variant="warning" className="mt-2 mb-0 small">
              <ul className="mb-0">{issues.map((s, i) => <li key={i}>{s}</li>)}</ul>
            </Alert>
          )}
          {notes.length > 0 && (
            <Alert variant="info" className="mt-2 mb-0 small">
              <ul className="mb-0">{notes.map((s, i) => <li key={i}>{s}</li>)}</ul>
            </Alert>
          )}
        </Card.Body>
      </Card>
    )
  }

  const workArea = (
    <div style={{ display: showAdmin ? 'none' : undefined }}>
      {error && <Alert variant="danger" dismissible onClose={() => setError(null)}>{error}</Alert>}

      <Row className="g-3 mb-3">
        {FIELDS.map(fd => (
          <Col md={4} key={fd.key}>
            <Card className="h-100">
              <Card.Header className={`bg-${fd.variant} text-white`}>{fd.title}</Card.Header>
              <Card.Body>
                <div className="text-muted small mb-2">{fd.hint}</div>
                <Form.Control type="file" multiple size="sm" accept=".docx,.docm,.pdf,.xlsx"
                              disabled={uploading[fd.key] || job?.running}
                              onChange={(e) => { uploadToField(fd.key, e.target.files); e.target.value = '' }} />
                <div className="d-flex justify-content-between align-items-center mt-1">
                  {uploading[fd.key] ? <span className="small text-muted"><Spinner size="sm" className="me-1" />Загрузка…</span> : <span />}
                  <Button variant="link" size="sm" className="p-0" disabled={job?.running}
                          onClick={() => uploadFolderToField(fd.key)}>из папки на сервере</Button>
                </div>
                <div className="mt-2">{filesOf(fd.key).map(renderFile)}</div>
              </Card.Body>
            </Card>
          </Col>
        ))}
      </Row>

      {analyzing && (
        <Alert variant="secondary" className="py-2 small">
          <Spinner size="sm" className="me-2" />
          {analysisProgress ? `${analysisProgress.message} · ${analysisProgress.detail || ''}` : 'Разбираю пакет…'}
        </Alert>
      )}

      {analysis && analysis.warnings.length > 0 && (
        <Alert variant="warning" className="small">
          <strong>Проверьте пакет:</strong>
          <ul className="mb-0 mt-1">{analysis.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul>
        </Alert>
      )}

      {analysis && analysis.checklists.length > 0 && (
        <Card className="mb-3">
          <Card.Header>✅ Что проверить</Card.Header>
          <Card.Body className="p-0">
            <Table size="sm" className="mb-0 align-middle">
              <thead className="table-light">
                <tr><th style={{ width: '3%' }}></th><th>Чек-лист</th><th>Проверяемый документ</th><th>Как проверяется</th></tr>
              </thead>
              <tbody>
                {analysis.checklists.map(c => (
                  <tr key={c.file} className={c.ready ? '' : 'text-muted'}>
                    <td><Form.Check type="checkbox" checked={!!selected[c.file]} disabled={!c.ready || job?.running}
                                    onChange={(e) => setSelected(prev => ({ ...prev, [c.file]: e.target.checked }))}
                                    aria-label={c.title} /></td>
                    <td><strong>{c.title}</strong><div className="small text-muted">{c.file}</div></td>
                    <td className="small">
                      {!c.ready && <span className="text-danger">{c.problem}</span>}
                      {c.ready && !c.multi && c.candidates.length > 1 && (
                        <Form.Select size="sm" value={c.selected[0]} disabled={job?.running}
                                     onChange={(e) => changeChoice(c.file, [e.target.value])}>
                          {c.candidates.map(n => <option key={n} value={n}>{n}</option>)}
                        </Form.Select>
                      )}
                      {c.ready && (c.multi || c.candidates.length === 1) && c.selected.join(', ')}
                      {c.warnings.filter(w => !w.startsWith('подходит несколько')).map((w, i) => (
                        <div key={i} className="text-warning-emphasis mt-1">⚠ {w}</div>
                      ))}
                    </td>
                    <td className="small">
                      {c.needs_gigachat ? <Badge bg="info">GigaChat · минуты</Badge> : <Badge bg="secondary">по формулам · секунды</Badge>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </Table>
          </Card.Body>
          <Card.Footer className="d-flex flex-wrap align-items-center gap-3">
            <Button variant="primary" onClick={startRun}
                    disabled={job?.running || analyzing || checklistsToRun.length === 0 || needsKey}>
              {job?.running ? <><Spinner size="sm" className="me-2" />Идёт проверка…</>
                : `🚀 Проверить выбранные (${checklistsToRun.length})`}
            </Button>
            {needsKey && <span className="small text-danger">Для Плана и Сводного акта сохраните ключ GigaChat в «Админ».</span>}
          </Card.Footer>
        </Card>
      )}

      {job?.running && (
        <Card className="mb-3">
          <Card.Body>
            <div className="d-flex justify-content-between mb-1">
              <strong>{job.checklist ? `${job.checklist}: ` : ''}{job.message}</strong>
              <span className="text-muted">⏱ {formatElapsed(elapsed)}</span>
            </div>
            <ProgressBar animated striped now={job.total ? (job.current / job.total) * 100 : 0} />
            {job.detail && <div className="small text-muted mt-1">{job.detail}</div>}
          </Card.Body>
        </Card>
      )}

      {job && !job.running && job.stage === 'error' && <Alert variant="danger">{job.message}</Alert>}

      {outputs.length > 0 && (
        <>
          <div className="d-flex align-items-center gap-2 mb-2">
            <h5 className="mb-0">Результаты</h5>
            {!job?.running && job?.message && <span className="text-muted small">{job.message}</span>}
            {outputs.filter(o => !o.error).length > 1 && !job?.running && (
              <Button size="sm" variant="outline-success" className="ms-auto"
                      onClick={() => window.open(`${API_BASE}/api/download-outputs/${sessionId}`, '_blank')}>
                🗜 Скачать всё (ZIP)
              </Button>
            )}
          </div>
          {outputs.map(renderOutput)}
        </>
      )}
    </div>
  )

  return (
    <Container fluid className="py-4">
      <div className="d-flex justify-content-end gap-2 mb-2">
        <Button size="sm" variant="outline-primary" onClick={() => setShowInstructions(true)}>
          📖 Инструкция
        </Button>
        {!showAdmin && sessionId && (
          <Button size="sm" variant="outline-secondary" onClick={newPackage} disabled={job?.running}>
            ＋ Новый пакет
          </Button>
        )}
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
          <Modal.Title>📖 Инструкция</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          <p>
            Загрузите пакет аудита в три поля — система сама определит, что за бланки и документы
            загружены, сопоставит каждый проверяемый документ с его чек-листом и покажет, что можно проверить.
          </p>
          <Card className="mb-3 border-primary">
            <Card.Header className="bg-primary text-white">📄 Поле 1 — бланки чек-листов</Card.Header>
            <Card.Body className="small">
              Пустые бланки «Чек лист …», «ИИ -План АУДИТА», «ЧК -Сводный акт» — сколько нужно. У каждого
              файла появится тип: «ЧК: Акт», «ЧК: ПН» и т. д. Бланк, который не распознан, подсвечивается красным.
            </Card.Body>
          </Card>
          <Card className="mb-3 border-info">
            <Card.Header className="bg-info text-white">📋 Поле 2 — проверяемые документы</Card.Header>
            <Card.Body className="small">
              <p className="mb-2">Документ, который проверяет каждый чек-лист:</p>
              <ul className="mb-2">
                <li>ЧК «План АУДИТА» — План аудита; ЧК «Сводный акт» — Сводный акт;</li>
                <li>ЧК «Акт» — Акт по результатам аудита; ЧК «Заключение» — Заключение об устранении НС;</li>
                <li>ЧК «Отчет 1 этапа» и «Заключение — 1 этап» — Отчет и Заключение 1 этапа;</li>
                <li>ЧК ПКД, Отчет КД, ПН, Листы оценки — План КД, Отчет по КД, Протоколы НС, Листы оценки.</li>
              </ul>
              У каждого документа видно, к какому чек-листу он пойдёт. Если подходящих документов несколько
              (например, Акт текущего и прошлого аудита), выберите нужный в таблице «Что проверить».
            </Card.Body>
          </Card>
          <Card className="mb-3 border-success">
            <Card.Header className="bg-success text-white">📚 Поле 3 — источники</Card.Header>
            <Card.Body className="small">
              Всё, с чем сверяются проверяемые документы: Договор и доп. соглашения, Приказ об ЭГ, Разбивка
              ОКВЭД, Информация для подготовки, Сертификат, Расчёт трудоёмкости, Акт (для ПКД, ПН, Заключения
              и Листов оценки), Отчет об устранении НС и т. д.
            </Card.Body>
          </Card>
          <Card className="mb-3 border-warning">
            <Card.Header className="bg-warning">⚠️ Тип документа</Card.Header>
            <Card.Body className="small">
              Тип определяется по содержимому (заголовку документа). У скана без текста — по имени файла:
              такой файл помечен «по имени — проверьте». Если тип неверный или «не определён», выберите
              его в списке под файлом — иначе документ не попадёт в проверку. Документы разных организаций
              и дубликаты система покажет в блоке «Проверьте пакет».
            </Card.Body>
          </Card>
          <Alert variant="secondary" className="mb-0 small">
            <strong>Порядок:</strong> бланки в поле 1 → документы в поля 2 и 3 → отметьте чек-листы в
            «Что проверить» → «Проверить выбранные». Чек-листы «по формулам» проверяются за секунды, План
            и Сводный акт — через GigaChat, за минуты (нужен ключ в «Админ»). Результаты появляются по мере
            готовности; пункты «вручную» подсвечены — в файле их клетки ОК/NOK остаются пустыми.
          </Alert>
        </Modal.Body>
        <Modal.Footer>
          <Button variant="primary" onClick={() => setShowInstructions(false)}>Понятно</Button>
        </Modal.Footer>
      </Modal>

      <h1 className="mb-4 text-center">
        <Badge bg="primary">Проверка документов аудита</Badge>
        <br />
        <span className="fs-5">Чек-листы ОС СМК: План, Акт, Заключения, ПКД, ПН, Сводный акт и др.</span>
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
            <Row className="g-2 align-items-center mt-1">
              <Col md={5}>
                <Form.Label className="small text-muted mb-1">
                  Тип аккаунта (OAuth scope)
                </Form.Label>
                <Form.Select
                  value={scope || ''}
                  onChange={(e) => { setScope(e.target.value); setApiKeySaved(false); setSaveSuccess(false) }}
                >
                  <option value="">— по умолчанию —</option>
                  {KNOWN_SCOPES.map((s) => (
                    <option key={s.value} value={s.value}>{s.label}</option>
                  ))}
                </Form.Select>
              </Col>
              <Col md={7}>
                <Form.Label className="small text-muted mb-1">&nbsp;</Form.Label>
                <div className="small text-muted">
                  Физлицо — PERS; организация/ИП — B2B (предоплата) или CORP (оплата по факту).
                  Должен совпадать с типом аккаунта, на который выпущен ключ.
                </div>
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

      {workArea}
    </Container>
  )
}

export default App
