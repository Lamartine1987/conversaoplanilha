import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Download, Settings, Plus, Trash2, Leaf, GitMerge, Search, Loader2, AlertCircle } from 'lucide-react';
import { collection, addDoc, getDocs, deleteDoc, doc, onSnapshot } from 'firebase/firestore';
import { db } from './firebase';

const Autocomplete = ({ options, value, onChange, placeholder }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const wrapperRef = React.useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [wrapperRef]);

  useEffect(() => {
    setSearch(value || '');
  }, [value]);

  const filteredOptions = options
    .filter(opt => opt.toLowerCase().includes(search.toLowerCase()))
    .slice(0, 100);

  return (
    <div ref={wrapperRef} style={{ position: 'relative', flex: 1 }}>
      <input
        type="text"
        value={search}
        onChange={(e) => {
          setSearch(e.target.value);
          onChange(e.target.value);
          setIsOpen(true);
        }}
        onFocus={() => setIsOpen(true)}
        placeholder={placeholder}
        style={{
          width: '100%',
          padding: '0.6rem 1rem',
          borderRadius: '6px',
          border: '1px solid #cbd5e1',
          fontSize: '0.9rem',
          outline: 'none',
          boxShadow: 'inset 0 1px 2px rgba(0,0,0,0.05)',
          color: '#334155'
        }}
      />
      {isOpen && (
        <ul style={{
          position: 'absolute',
          top: 'calc(100% + 4px)',
          left: 0,
          right: 0,
          maxHeight: '220px',
          overflowY: 'auto',
          backgroundColor: '#fff',
          border: '1px solid #cbd5e1',
          borderRadius: '6px',
          boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05)',
          zIndex: 999,
          margin: 0,
          padding: '0.25rem 0',
          listStyle: 'none'
        }}>
          {filteredOptions.length > 0 ? filteredOptions.map(opt => (
            <li 
              key={opt}
              onClick={() => {
                onChange(opt);
                setSearch(opt);
                setIsOpen(false);
              }}
              style={{
                padding: '0.6rem 1rem',
                cursor: 'pointer',
                fontSize: '0.9rem',
                color: '#334155',
                backgroundColor: '#fff'
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f5f9'}
              onMouseLeave={(e) => e.target.style.backgroundColor = '#fff'}
            >
              {opt}
            </li>
          )) : (
            <li style={{ padding: '0.6rem 1rem', color: '#94a3b8', fontSize: '0.9rem' }}>Nenhum produto encontrado.</li>
          )}
        </ul>
      )}
    </div>
  );
};

function App() {
  const [oldFile, setOldFile] = useState(null);
  const [newFile, setNewFile] = useState(null);
  
  const [oldHeaders, setOldHeaders] = useState([]);
  const [newHeaders, setNewHeaders] = useState([]);

  const [mapKeyOld, setMapKeyOld] = useState('');
  const [mapKeyNew, setMapKeyNew] = useState('');
  const [dataMappings, setDataMappings] = useState([{ oldCol: '', newCol: '' }]);

  const [rules, setRules] = useState([]);
  const [ruleOld, setRuleOld] = useState('');
  const [ruleNew, setRuleNew] = useState('');
  const [processing, setProcessing] = useState(false);
  const [message, setMessage] = useState({ type: '', text: '' });
  const [unmatchedItems, setUnmatchedItems] = useState([]);
  const [oldKeys, setOldKeys] = useState([]);
  const [uncheckedColumns, setUncheckedColumns] = useState([]);
  const [searchUnmatched, setSearchUnmatched] = useState('');
  const [selectedOldKeys, setSelectedOldKeys] = useState({});

  // States for Analysis and Generation
  const [analysisComplete, setAnalysisComplete] = useState(false);
  const [analysisStats, setAnalysisStats] = useState({ matched: 0, unmatched: 0, blank: 0, total: 0 });
  const [finalData, setFinalData] = useState([]);
  const [parsedOld, setParsedOld] = useState(null);
  const [parsedNew, setParsedNew] = useState(null);
  const [parsedWorkbook, setParsedWorkbook] = useState(null);
  const [parsedSheetName, setParsedSheetName] = useState('');
  
  // Stats
  const [oldDataCount, setOldDataCount] = useState(0);
  const [newDataCount, setNewDataCount] = useState(0);

  // Load Rules from Firestore
  useEffect(() => {
    const rulesCollection = collection(db, 'regras_mapeamento');
    const unsubscribe = onSnapshot(rulesCollection, (snapshot) => {
      const rulesList = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setRules(rulesList);
    }, (error) => {
      console.error("Erro ao carregar regras do Firebase:", error);
      setMessage({ type: 'warning', text: 'Aviso: Firebase não configurado corretamente. As regras não serão salvas. Configure o .env' });
    });

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (analysisComplete) {
      analyzeFiles(true);
    }
  }, [rules, uncheckedColumns]);

  const extractHeadersAndData = (file, type) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0] || [];
      const cleanHeaders = headers.map(h => String(h).trim());
      
      const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      
      const json = rawData.map(row => {
        const cleanRow = {};
        Object.keys(row).forEach(key => {
          cleanRow[key.trim()] = row[key];
        });
        return cleanRow;
      });

      console.log(`[Extrator] Arquivo tipo: ${type} | Total de Linhas Lidas: ${json.length} | Colunas: ${cleanHeaders.length}`);

      if (type === 'old') {
        setOldHeaders(cleanHeaders);
        setParsedOld(json);
        setOldDataCount(json.length);
      } else {
        setNewHeaders(cleanHeaders);
        setParsedNew(json);
        setParsedWorkbook(workbook);
        setParsedSheetName(sheetName);
        setNewDataCount(json.length);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileUpload = (e, setFile, type) => {
    const file = e.target.files[0];
    if (file) {
      setFile(file);
      extractHeadersAndData(file, type);
      setMessage({ type: '', text: '' });
      setAnalysisComplete(false);
    }
    // Reseta o input para permitir selecionar o mesmo arquivo novamente
    e.target.value = null;
  };

  const clearFile = (type) => {
    if (type === 'old') {
      setOldFile(null);
      setOldHeaders([]);
      setParsedOld(null);
      setOldDataCount(0);
      setMapKeyOld('');
    } else {
      setNewFile(null);
      setNewHeaders([]);
      setParsedNew(null);
      setParsedWorkbook(null);
      setParsedSheetName('');
      setNewDataCount(0);
      setMapKeyNew('');
    }
    setAnalysisComplete(false);
    setMessage({ type: '', text: '' });
  };

  const addRule = async (e) => {
    e.preventDefault();
    if (!ruleOld || !ruleNew) return;
    
    try {
      await addDoc(collection(db, 'regras_mapeamento'), {
        oldKeyword: ruleOld,
        newKeyword: ruleNew,
        createdAt: new Date()
      });
      setRuleOld('');
      setRuleNew('');
    } catch (error) {
      console.error("Erro ao adicionar regra:", error);
      setMessage({ type: 'error', text: 'Erro ao salvar regra. Verifique a configuração do Firebase.' });
    }
  };

  const addRuleDirectly = async (oldName, newName) => {
    if (!oldName) {
      setMessage({ type: 'error', text: 'Selecione um produto original para vincular.' });
      return;
    }
    
    if (!oldKeys.includes(oldName)) {
      setMessage({ type: 'error', text: `O produto "${oldName}" não existe na Planilha Antiga. Por favor, escolha uma opção válida da lista.` });
      return;
    }

    try { 
      await addDoc(collection(db, 'regras_mapeamento'), {
        oldKeyword: oldName,
        newKeyword: newName,
        createdAt: new Date()
      });
      setMessage({ type: 'success', text: `Vínculo criado com sucesso para ${newName}!` });
      // Limpa o input específico após sucesso
      setSelectedOldKeys(prev => {
        const next = { ...prev };
        delete next[newName];
        return next;
      });
    } catch (error) {
      console.error("Erro ao adicionar regra:", error);
      setMessage({ type: 'error', text: 'Erro ao salvar vínculo.' });
    }
  };

  const deleteRule = async (id) => {
    try {
      await deleteDoc(doc(db, 'regras_mapeamento', id));
    } catch (error) {
      console.error("Erro ao excluir regra:", error);
    }
  };

  const analyzeFiles = async (isSilent = false) => {
    if (!oldFile || !newFile || !parsedOld || !parsedNew) {
      if (!isSilent) setMessage({ type: 'error', text: 'Por favor, aguarde o carregamento ou selecione as planilhas.' });
      return;
    }

    if (!mapKeyOld || !mapKeyNew || dataMappings.some(m => !m.oldCol || !m.newCol)) {
      if (!isSilent) setMessage({ type: 'error', text: 'Preencha todos os mapeamentos antes de analisar.' });
      return;
    }

    if (!isSilent) {
      setProcessing(true);
      setMessage({ type: 'info', text: 'Analisando correspondências...' });
      
      await new Promise(resolve => setTimeout(resolve, 50));
    }

    try {
      const oldData = parsedOld;
      const newData = parsedNew;

      let matchedCount = 0;
      let blankCount = 0;

      const normalizeStr = (str) => {
        return String(str || '').replace(/\s+/g, ' ').trim().toUpperCase();
      };

      const rulesMap = {};
      rules.forEach(r => {
        rulesMap[normalizeStr(r.oldKeyword)] = normalizeStr(r.newKeyword);
      });

      const oldDataMap = {};
      const uniqueOldKeys = new Set();
      
      oldData.forEach(row => {
        const originalOldKey = String(row[mapKeyOld] || '').trim();
        const oldKeyVal = normalizeStr(originalOldKey);
        if (!oldKeyVal) return;

        uniqueOldKeys.add(originalOldKey);
        let mappedKey = oldKeyVal;

        if (rulesMap[oldKeyVal]) {
          mappedKey = rulesMap[oldKeyVal];
        }

        const extractedData = {};
        dataMappings.forEach(mapping => {
          extractedData[mapping.oldCol] = row[mapping.oldCol];
        });

        oldDataMap[mappedKey] = extractedData;
      });

      const unmatchedList = [];

      const formatNumber = (val) => {
        if (val === null || val === undefined || val === '') return '';
        if (typeof val === 'number') return val;

        const strVal = String(val).trim();

        // Preserva zeros à esquerda (CPFs, CNPJs, Códigos)
        if (/^0\d+$/.test(strVal)) {
          return strVal;
        }

        // Tenta converter formatos brasileiros (ex: 1.234,56 ou 196,27) para número real
        if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(strVal) || /^-?\d+(,\d+)?$/.test(strVal)) {
          const cleanStr = strVal.replace(/\./g, '').replace(',', '.');
          const num = parseFloat(cleanStr);
          if (!isNaN(num)) return num;
        }

        return strVal;
      };

      const allOutputColumns = Array.from(new Set([...newHeaders, ...dataMappings.map(m => m.newCol).filter(Boolean)]));
      const finalOutputColumns = allOutputColumns.filter(c => !uncheckedColumns.includes(c));

      const updatedNewData = newData.map(row => {
        const newKeyVal = normalizeStr(row[mapKeyNew]);
        const match = oldDataMap[newKeyVal];

        const updatedRow = { ...row };

        if (match) {
          matchedCount++;
          dataMappings.forEach(mapping => {
            updatedRow[mapping.newCol] = formatNumber(match[mapping.oldCol]);
          });
        } else if (newKeyVal) {
          unmatchedList.push({ name: newKeyVal });
          dataMappings.forEach(mapping => {
            updatedRow[mapping.newCol] = '';
          });
        } else {
          blankCount++;
        }

        const finalRow = {};
        finalOutputColumns.forEach(col => {
          finalRow[col] = updatedRow[col] !== undefined ? updatedRow[col] : '';
        });

        return finalRow;
      });

      setUnmatchedItems(unmatchedList);
      setOldKeys(Array.from(uniqueOldKeys).sort());
      setFinalData(updatedNewData);
      
      setAnalysisStats({
        matched: matchedCount,
        unmatched: unmatchedList.length,
        blank: blankCount,
        total: newData.length
      });
      setAnalysisComplete(true);
      
      if (!isSilent) {
        setMessage({ type: 'success', text: `Análise concluída! ${matchedCount} produtos correspondidos. Verifique as estatísticas abaixo.` });
      }
    } catch (error) {
      console.error(error);
      if (!isSilent) setMessage({ type: 'error', text: 'Ocorreu um erro ao analisar os dados.' });
    } finally {
      if (!isSilent) setProcessing(false);
    }
  };

  const downloadExcel = () => {
    if (!analysisComplete || !parsedWorkbook) return;
    const newWorksheet = XLSX.utils.json_to_sheet(finalData);

    if (newWorksheet['!ref']) {
      const range = XLSX.utils.decode_range(newWorksheet['!ref']);
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const headerCell = newWorksheet[XLSX.utils.encode_cell({r: 0, c: C})];
        if (headerCell && headerCell.v) {
          if (dataMappings.some(m => m.newCol === headerCell.v)) {
            for (let R = 1; R <= range.e.r; ++R) {
              const cell = newWorksheet[XLSX.utils.encode_cell({r: R, c: C})];
              if (cell && typeof cell.v === 'number') {
                // Aplica formatação decimal apenas se o número tiver casas decimais (não for inteiro)
                if (!Number.isInteger(cell.v)) {
                  cell.z = '#,##0.00';
                }
              }
            }
          }
        }
      }
    }

    parsedWorkbook.Sheets[parsedSheetName] = newWorksheet;
    XLSX.writeFile(parsedWorkbook, `Final_${newFile.name}`);
    setMessage({ type: 'success', text: `Sucesso! Planilha Final gerada e baixada.` });
  };

  return (
    <div className="container">
      <header className="app-header">
        <Leaf size={40} color="var(--color-primary)" />
        <h1 className="app-title">AgroSync - Conversor Universal</h1>
      </header>

      {message.text && (
        <div className={`alert alert-${message.type}`}>
          {message.text}
        </div>
      )}

      <div className="grid-2">
        <div className="card">
          <div className="card-title">
            <FileSpreadsheet />
            Planilha Antiga (Origem)
          </div>
          <p style={{marginBottom: '1rem', color: 'var(--color-text-muted)', fontSize: '0.9rem'}}>
            A planilha que contém os dados que você quer extrair.
          </p>
          <label className="upload-area">
            <Upload className="upload-icon" size={32} />
            <div>Clique ou arraste (.xlsx, .csv)</div>
            <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => handleFileUpload(e, setOldFile, 'old')} style={{ display: 'none' }} />
          </label>
          {oldFile && (
             <div className="file-info" style={{marginTop: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', padding: '0.8rem 1rem', borderRadius: '8px', border: '1px solid #e2e8f0'}}>
               <div>
                 <strong>Arquivo:</strong> {oldFile.name}<br/>
                 <span style={{color: '#64748b', fontSize: '0.85rem'}}>O sistema identificou <strong>{oldDataCount}</strong> linhas com dados neste arquivo.</span>
               </div>
               <button 
                 className="btn btn-danger" 
                 style={{padding: '0.4rem 0.8rem', fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: '0.3rem'}}
                 onClick={() => clearFile('old')}
                 title="Remover planilha"
               >
                 <Trash2 size={16} /> Remover
               </button>
             </div>
          )}
        </div>

        <div className="card">
          <div className="card-title">
            <FileSpreadsheet />
            Planilha Nova (Destino)
          </div>
          <p style={{marginBottom: '1rem', color: 'var(--color-text-muted)', fontSize: '0.9rem'}}>
            A planilha que receberá os dados extraídos.
          </p>
          <label className="upload-area">
            <Upload className="upload-icon" size={32} />
            <div>Clique ou arraste (.xlsx, .csv)</div>
            <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => handleFileUpload(e, setNewFile, 'new')} style={{ display: 'none' }} />
          </label>
          {newFile && (
             <div className="file-info" style={{marginTop: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', padding: '0.8rem 1rem', borderRadius: '8px', border: '1px solid #e2e8f0'}}>
               <div>
                 <strong>Arquivo:</strong> {newFile.name}<br/>
                 <span style={{color: '#64748b', fontSize: '0.85rem'}}>O sistema identificou <strong>{newDataCount}</strong> linhas com dados neste arquivo.</span>
               </div>
               <button 
                 className="btn btn-danger" 
                 style={{padding: '0.4rem 0.8rem', fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: '0.3rem'}}
                 onClick={() => clearFile('new')}
                 title="Remover planilha"
               >
                 <Trash2 size={16} /> Remover
               </button>
             </div>
          )}
        </div>
      </div>

      {oldHeaders.length > 0 && newHeaders.length > 0 && (
        <div className="card" style={{borderColor: 'var(--color-primary-light)', borderWidth: '2px'}}>
          <div className="card-title">
            <GitMerge /> Mapeamento de Colunas
          </div>
          <p style={{color: 'var(--color-text-muted)', marginBottom: '1.5rem'}}>
            Configure como o sistema deve conectar as duas planilhas.
          </p>
          
          <div style={{marginBottom: '1.5rem', padding: '1.5rem', background: 'var(--color-bg)', borderRadius: 'var(--radius-md)', border: '1px solid var(--color-border)'}}>
            <h4 style={{marginBottom: '1rem', fontSize: '1.1rem'}}>1. Chave de Ligação (Identificador Principal)</h4>
            <p style={{fontSize: '0.9rem', color: '#64748b', marginBottom: '1rem'}}>
              Escolha a coluna em cada planilha que contém a informação em comum (ex: Nome do Cliente, Código, CPF) que será usada para "casar" as linhas.
            </p>
            <div className="grid-2">
              <div>
                <label className="input-label">Coluna na Antiga:</label>
                <select className="input-field" style={{width: '100%', marginTop: '0.5rem'}} value={mapKeyOld} onChange={e => setMapKeyOld(e.target.value)}>
                  <option value="">Selecione a coluna...</option>
                  {oldHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="input-label">Coluna na Nova:</label>
                <select className="input-field" style={{width: '100%', marginTop: '0.5rem'}} value={mapKeyNew} onChange={e => setMapKeyNew(e.target.value)}>
                  <option value="">Selecione a coluna...</option>
                  {newHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>

          <div style={{padding: '1.5rem', background: 'var(--color-bg)', borderRadius: 'var(--radius-md)', border: '1px solid var(--color-border)'}}>
            <h4 style={{marginBottom: '1rem', fontSize: '1.1rem'}}>2. Dados a Preencher (Copiar e Colar)</h4>
            <p style={{fontSize: '0.9rem', color: '#64748b', marginBottom: '1rem'}}>
              Escolha quais colunas você quer "puxar" da planilha antiga. Você pode selecionar uma coluna existente da nova para ser substituída, <strong>ou digitar um novo nome</strong> (ex: "Valores a Pagar") para criar uma coluna nova!
            </p>
            {dataMappings.map((mapping, index) => (
              <div key={index} className="grid-2" style={{marginBottom: '1rem', alignItems: 'flex-end'}}>
                <div>
                  <label className="input-label">Copiar desta coluna (Antiga):</label>
                  <select className="input-field" style={{width: '100%', marginTop: '0.5rem'}} value={mapping.oldCol} onChange={e => {
                    const newMappings = [...dataMappings];
                    newMappings[index].oldCol = e.target.value;
                    setDataMappings(newMappings);
                  }}>
                    <option value="">Selecione o dado...</option>
                    {oldHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
                <div style={{display: 'flex', gap: '0.5rem'}}>
                  <div style={{flex: 1}}>
                    <label className="input-label">Nome da coluna na Nova:</label>
                    <input 
                      list="new-headers-list"
                      className="input-field" 
                      style={{width: '100%', marginTop: '0.5rem'}} 
                      value={mapping.newCol} 
                      onChange={e => {
                        const newMappings = [...dataMappings];
                        newMappings[index].newCol = e.target.value;
                        setDataMappings(newMappings);
                      }}
                      placeholder="Ex: PREÇO NOVO"
                    />
                    <datalist id="new-headers-list">
                      {newHeaders.map(h => <option key={h} value={h} />)}
                    </datalist>
                  </div>
                  <button 
                    className="btn btn-danger" 
                    style={{height: '42px'}}
                    disabled={dataMappings.length === 1}
                    onClick={() => setDataMappings(dataMappings.filter((_, i) => i !== index))}
                  >
                    <Trash2 size={20} />
                  </button>
                </div>
              </div>
            ))}
            <button className="btn btn-outline" style={{marginTop: '0.5rem'}} onClick={() => setDataMappings([...dataMappings, {oldCol: '', newCol: ''}])}>
              <Plus size={18} /> Adicionar outro dado
            </button>
          </div>

          <div style={{padding: '1.5rem', background: 'var(--color-bg)', borderRadius: 'var(--radius-md)', border: '1px solid var(--color-border)', marginTop: '1.5rem'}}>
            <h4 style={{marginBottom: '1rem', fontSize: '1.1rem'}}>3. Limpeza do Arquivo Final</h4>
            <p style={{fontSize: '0.9rem', color: '#64748b', marginBottom: '1rem'}}>
              Desmarque as colunas que você <strong>não</strong> quer que apareçam no arquivo final gerado.
            </p>
            <div style={{display: 'flex', flexWrap: 'wrap', gap: '0.8rem'}}>
              {Array.from(new Set([...newHeaders, ...dataMappings.map(m => m.newCol).filter(Boolean)])).map(col => {
                const isChecked = !uncheckedColumns.includes(col);
                return (
                  <label key={col} style={{display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer', background: isChecked ? '#e0f2fe' : '#f1f5f9', padding: '0.4rem 0.8rem', borderRadius: '20px', border: `1px solid ${isChecked ? '#7dd3fc' : '#cbd5e1'}`}}>
                    <input 
                      type="checkbox" 
                      checked={isChecked} 
                      onChange={() => {
                        if (isChecked) {
                          setUncheckedColumns([...uncheckedColumns, col]);
                        } else {
                          setUncheckedColumns(uncheckedColumns.filter(c => c !== col));
                        }
                      }} 
                    />
                    <span style={{color: isChecked ? '#0369a1' : '#64748b', fontWeight: isChecked ? 500 : 400, fontSize: '0.9rem'}}>{col}</span>
                  </label>
                )
              })}
            </div>
          </div>
        </div>
      )}

      <div className="card">
        <div className="card-title">
          <Settings />
          Regras de Mapeamento Manual (Dicionário)
        </div>
        <p style={{marginBottom: '1.5rem', color: '#64748b', fontSize: '0.9rem'}}>
          O sistema usará a <strong>Chave de Ligação</strong> escolhida. Se a chave original estiver escrita diferente da nova (ex: "BOTA INF" e "BOTA INFANTIL"), você pode ensinar ao sistema criar uma regra de tradução aqui:
        </p>
        <form className="rule-form" onSubmit={addRule}>
          <div className="input-group">
            <label className="input-label">Se o identificador original contiver:</label>
            <input type="text" className="input-field" value={ruleOld} onChange={e => setRuleOld(e.target.value)} placeholder="Ex: ABOR-VAC" />
          </div>
          <div className="input-group">
            <label className="input-label">O identificador na nova será:</label>
            <input type="text" className="input-field" value={ruleNew} onChange={e => setRuleNew(e.target.value)} placeholder="Ex: ABORVAC" />
          </div>
          <button type="submit" className="btn btn-primary" disabled={!ruleOld || !ruleNew}>
            <Plus size={18} /> Adicionar Regra
          </button>
        </form>

        {rules.length > 0 ? (
          <table className="rules-table">
            <thead>
              <tr>
                <th>Se Original contém...</th>
                <th>Identificador Novo correspondente...</th>
                <th style={{width: '80px'}}>Ação</th>
              </tr>
            </thead>
            <tbody>
              {rules.map(rule => (
                <tr key={rule.id}>
                  <td><strong>{rule.oldKeyword}</strong></td>
                  <td><span style={{color: 'var(--color-primary)', fontWeight: 500}}>{rule.newKeyword}</span></td>
                  <td>
                    <button onClick={() => deleteRule(rule.id)} className="btn btn-danger" title="Excluir regra">
                      <Trash2 size={18} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div style={{textAlign: 'center', padding: '2rem', color: 'var(--color-text-muted)'}}>
            Nenhuma regra cadastrada ainda.
          </div>
        )}
      </div>

      <div style={{textAlign: 'center', marginTop: '2rem'}}>
        <button 
          className="btn btn-primary" 
          style={{fontSize: '1.2rem', padding: '1rem 2rem'}} 
          onClick={() => {
            setProcessing(true);
            setTimeout(() => analyzeFiles(false), 100);
          }}
          disabled={processing || !oldFile || !newFile || !mapKeyOld || !mapKeyNew || dataMappings.some(m => !m.oldCol || !m.newCol)}
        >
          {processing ? <><Loader2 className="animate-spin" /> Analisando...</> : <><Search /> 1º Passo: Analisar Correspondências</>}
        </button>
      </div>

      {analysisComplete && unmatchedItems.length > 0 && (
          <div className="card fade-in" style={{ borderColor: '#fef08a', backgroundColor: '#fefce8' }}>
            <div className="card-header" style={{ borderBottom: '1px solid #fef08a' }}>
              <h2 className="card-title" style={{ color: '#b45309' }}>
                <AlertCircle size={20} style={{ marginRight: '8px' }} />
                Produtos Sem Correspondência ({unmatchedItems.length})
              </h2>
            </div>
            
            <div style={{ padding: '1rem', borderBottom: '1px solid #fef08a' }}>
            <div style={{display: 'flex', alignItems: 'center', background: '#fff', padding: '0.5rem 1rem', borderRadius: '8px', border: '1px solid #fde047'}}>
              <Search size={18} color="#a16207" style={{marginRight: '8px'}} />
              <input 
                type="text" 
                placeholder="Buscar produto pendente..." 
                value={searchUnmatched}
                onChange={(e) => setSearchUnmatched(e.target.value)}
                style={{border: 'none', outline: 'none', width: '100%', background: 'transparent', color: '#a16207'}}
              />
            </div>
          </div>

          <div style={{maxHeight: '500px', minHeight: '350px', overflowY: 'auto', backgroundColor: '#fff', borderRadius: '8px', border: '1px solid #fef08a'}}>
            <table className="rules-table" style={{marginTop: 0}}>
              <thead style={{position: 'sticky', top: 0, zIndex: 1, backgroundColor: '#fef9c3'}}>
                <tr>
                  <th>Identificador (Planilha Nova)</th>
                  <th>Ação Rápida: Vincular à Planilha Antiga</th>
                </tr>
              </thead>
              <tbody>
                {unmatchedItems.filter(item => item.name.toLowerCase().includes(searchUnmatched.toLowerCase())).map((item, index) => (
                  <tr key={index}>
                    <td style={{width: '30%'}}><strong>{item.name}</strong></td>
                    <td style={{position: 'relative'}}>
                      <div style={{display: 'flex', gap: '0.5rem', alignItems: 'center'}}>
                        <Autocomplete
                          options={oldKeys}
                          value={selectedOldKeys[item.name] || ''}
                          onChange={(val) => setSelectedOldKeys({ ...selectedOldKeys, [item.name]: val })}
                          placeholder="Digite para buscar na planilha antiga..."
                        />
                        <button 
                          className="btn btn-primary"
                          style={{padding: '0.6rem 1rem', fontSize: '0.9rem', whiteSpace: 'nowrap'}}
                          onClick={() => addRuleDirectly(selectedOldKeys[item.name], item.name)}
                        >
                          <GitMerge size={14} style={{marginRight: '4px'}}/> Vincular
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {analysisComplete && (
        <div className="card" style={{marginTop: '2rem', textAlign: 'center', borderColor: '#10B981', borderWidth: '2px', backgroundColor: '#ecfdf5', padding: '3rem 2rem'}}>
          <h3 style={{color: '#065f46', marginBottom: '1rem', fontSize: '1.5rem'}}>Tudo pronto para exportar!</h3>
          <p style={{color: '#064e3b', marginBottom: '2rem', fontSize: '1.1rem'}}>
             A planilha foi analisada e os vínculos foram aplicados. Se as pendências acima estiverem resolvidas ou se você já quiser baixar, clique no botão abaixo.
          </p>
          <button 
            className="btn btn-primary" 
            style={{fontSize: '1.3rem', padding: '1rem 3rem', background: '#10B981', boxShadow: '0 10px 15px -3px rgba(16, 185, 129, 0.3)'}} 
            onClick={downloadExcel}
          >
            <Download /> 2º Passo: Gerar e Baixar Planilha Final
          </button>
        </div>
      )}
    </div>
  );
}

export default App;
