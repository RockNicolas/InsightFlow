import { useCallback, useEffect, useRef, useState } from 'react'

import { downloadUrl, fetchSession, generateReports, uploadWorkbook } from '../api/client'
import { AppNav } from '../components/AppNav'
import { Messages } from '../components/Messages'
import { ReportCheckbox } from '../components/ReportCheckbox'
import type { GeneratePayload, SessionState } from '../types'

const emptySession: SessionState = {
  uploaded_name: '',
  fleet_profiles: [],
  selected_fleet_key: 'saneamento',
  selected_fleet_label: 'Saneamento',
  sheet_names: [],
  weekly_options: [],
  weekly_sheet: '',
  observation_sheet: '',
  generate_all_month_sheets: false,
  results: [],
  output_folder: '',
}

export function ReportsPage() {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [session, setSession] = useState<SessionState>(emptySession)
  const [fleetKey, setFleetKey] = useState('saneamento')
  const [weeklySheet, setWeeklySheet] = useState('')
  const [observationSheet, setObservationSheet] = useState('')
  const [fileLabel, setFileLabel] = useState('Nenhum arquivo selecionado')
  const [uploadHint, setUploadHint] = useState(
    'Ao escolher o arquivo, as abas serao carregadas automaticamente.',
  )
  const [messages, setMessages] = useState<Array<{ type: 'success' | 'error'; text: string }>>([])
  const [loading, setLoading] = useState(true)
  const [uploading, setUploading] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [options, setOptions] = useState({
    generate_weekly: false,
    generate_observation: false,
    generate_monthly: false,
    generate_all_month_sheets: false,
  })

  const pushMessage = useCallback((type: 'success' | 'error', text: string) => {
    setMessages([{ type, text }])
  }, [])

  const applySession = useCallback((data: SessionState) => {
    setSession(data)
    setFleetKey(data.selected_fleet_key)
    setWeeklySheet(data.weekly_sheet)
    setObservationSheet(data.observation_sheet)
    if (data.uploaded_name) {
      setFileLabel(data.uploaded_name)
    }
  }, [])

  useEffect(() => {
    fetchSession()
      .then((response) => {
        if (response.ok) {
          applySession(response.data)
        } else {
          pushMessage('error', response.message)
        }
      })
      .catch(() => pushMessage('error', 'Nao foi possivel carregar a sessao.'))
      .finally(() => setLoading(false))
  }, [applySession, pushMessage])

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) {
      setFileLabel('Nenhum arquivo selecionado')
      return
    }

    setFileLabel(file.name)
    setUploadHint('Carregando abas...')
    setUploading(true)
    setMessages([])

    try {
      const response = await uploadWorkbook(fleetKey, file)
      if (response.ok) {
        applySession(response.data)
        pushMessage('success', response.message)
      } else {
        pushMessage('error', response.message)
      }
    } catch {
      pushMessage('error', 'Falha ao enviar a planilha.')
    } finally {
      setUploading(false)
      setUploadHint('Ao escolher o arquivo, as abas serao carregadas automaticamente.')
      if (fileInputRef.current) {
        fileInputRef.current.value = ''
      }
    }
  }

  const handleGenerate = async (event: React.FormEvent) => {
    event.preventDefault()
    setGenerating(true)
    setMessages([])

    const payload: GeneratePayload = {
      weekly_sheet: weeklySheet,
      observation_sheet: observationSheet,
      ...options,
    }

    try {
      const response = await generateReports(payload)
      if (response.ok) {
        applySession(response.data)
        pushMessage('success', response.message)
      } else {
        pushMessage('error', response.message)
      }
    } catch {
      pushMessage('error', 'Falha ao gerar os relatórios.')
    } finally {
      setGenerating(false)
    }
  }

  if (loading) {
    return (
      <main className="page">
        <AppNav />
        <section className="card">
          <p>Carregando interface...</p>
        </section>
      </main>
    )
  }

  return (
    <main className="page">
      <AppNav />

      <section className="hero hero-with-action">
        <div>
          <span className="eyebrow">InsightFlow</span>
          <h1>Interface web para gerar relatórios</h1>
          <p>
            Envie a planilha Excel, visualize as abas encontradas e gere os relatórios
            semanal, observações e mensal em um layout mais amplo.
          </p>
        </div>
       
      </section>

      <Messages items={messages} />

      <section className="workspace">
        <section className="card upload-card">
          <div className="card-header">
            <h2>1. Carregar planilha</h2>
            <p>
              Escolha a frota correta e envie o Excel correspondente. Cada frota usa seu
              próprio fluxo de leitura.
            </p>
          </div>

          <div className="stack upload-form">
            <label>
              <span>Frota</span>
              <select
                value={fleetKey}
                onChange={(event) => setFleetKey(event.target.value)}
                disabled={uploading}
              >
                {session.fleet_profiles.map((profile) => (
                  <option key={profile.key} value={profile.key}>
                    {profile.label}
                  </option>
                ))}
              </select>
            </label>

            <div className="file-picker">
              <span>Escolher arquivo Excel</span>
              <div className="file-picker-row">
                <input
                  ref={fileInputRef}
                  id="excel-file"
                  className="file-input"
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  onChange={handleFileChange}
                  disabled={uploading}
                />
                <label htmlFor="excel-file" className="file-button">
                  {uploading ? 'Carregando...' : 'Escolher arquivo'}
                </label>
                <span className="file-name">{fileLabel}</span>
              </div>
              <small className="upload-hint">{uploadHint}</small>
            </div>
          </div>

          {session.uploaded_name ? (
            <div className="loaded-file">
              <span className="label">Arquivo atual</span>
              <strong>{session.selected_fleet_label}</strong>
              <strong>{session.uploaded_name}</strong>
            </div>
          ) : null}
        </section>

        <section className="card config-card">
          <div className="card-header">
            <h2>2. Escolher abas e relatórios</h2>
            <p>
              {session.selected_fleet_label} selecionada. O relatório mensal usa
              automaticamente todas as abas semanais do arquivo.
            </p>
          </div>

          {session.sheet_names.length > 0 ? (
            <>
              <div className="sheet-preview">
                <span className="label">Abas encontradas no arquivo</span>
                <div className="sheet-tags">
                  {session.sheet_names.map((sheet) => (
                    <span
                      key={sheet}
                      className={`sheet-tag${sheet === observationSheet ? ' observation' : ''}`}
                    >
                      {sheet}
                    </span>
                  ))}
                </div>
              </div>

              <form className="stack" onSubmit={handleGenerate}>
                <div className="form-layout">
                  <div className="form-main">
                    <div className="grid controls-grid">
                      <label>
                        <span>Aba semanal</span>
                        <select
                          value={weeklySheet}
                          onChange={(event) => setWeeklySheet(event.target.value)}
                          disabled={generating}
                        >
                          <option value="">Selecione a aba semanal</option>
                          {session.weekly_options.map((sheet) => (
                            <option key={sheet} value={sheet}>
                              {sheet}
                            </option>
                          ))}
                        </select>
                      </label>

                      <label>
                        <span>Aba de observações</span>
                        <select
                          value={observationSheet}
                          onChange={(event) => setObservationSheet(event.target.value)}
                          disabled={generating}
                        >
                          <option value="">Nenhuma</option>
                          {session.sheet_names.map((sheet) => (
                            <option key={sheet} value={sheet}>
                              {sheet}
                            </option>
                          ))}
                        </select>
                      </label>
                    </div>
                  </div>

                  <aside className="options-panel">
                    <span className="label">Relatórios</span>
                    <div className="options">
                      <ReportCheckbox
                        checked={options.generate_weekly}
                        onChange={(checked) =>
                          setOptions((current) => ({ ...current, generate_weekly: checked }))
                        }
                        title="Gerar relatório semanal"
                        description="Gera o PDF da aba semanal escolhida."
                      />
                      <ReportCheckbox
                        checked={options.generate_observation}
                        onChange={(checked) =>
                          setOptions((current) => ({
                            ...current,
                            generate_observation: checked,
                          }))
                        }
                        title="Gerar relatório de observações"
                        description="Cria o PDF com os apontamentos e anotações."
                      />
                      <ReportCheckbox
                        checked={options.generate_monthly}
                        onChange={(checked) =>
                          setOptions((current) => ({ ...current, generate_monthly: checked }))
                        }
                        title="Gerar relatório mensal"
                        description="Monta o resumo executivo do mês."
                      />
                      <ReportCheckbox
                        checked={options.generate_all_month_sheets}
                        onChange={(checked) =>
                          setOptions((current) => ({
                            ...current,
                            generate_all_month_sheets: checked,
                          }))
                        }
                        title="Baixar todas as abas do mês"
                        description="Gera um relatório semanal para cada aba encontrada."
                      />
                    </div>

                    <button type="submit" className="button primary" disabled={generating}>
                      {generating ? 'Gerando...' : 'Gerar relatórios'}
                    </button>
                  </aside>
                </div>
              </form>
            </>
          ) : (
            <div className="empty-state">
              <p>Carregue uma planilha para desbloquear a seleção das abas.</p>
            </div>
          )}
        </section>
      </section>

      {session.results.length > 0 ? (
        <section className="card">
          <div className="card-header">
            <h2>3. Arquivos gerados</h2>
            <p>Abra os PDFs diretamente pelo navegador.</p>
          </div>

          <div className="results">
            {session.results.map((result) => (
              <article key={result.download_path} className="result-card">
                <span>{result.label}</span>
                <strong>{result.filename}</strong>
                <a
                  href={downloadUrl(result.download_path)}
                  target="_blank"
                  rel="noreferrer"
                >
                  Abrir PDF
                </a>
              </article>
            ))}
          </div>
        </section>
      ) : null}
    </main>
  )
}
