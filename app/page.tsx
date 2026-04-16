'use client';

import React, { useState } from 'react';

export default function Home() {
  const [activeTab, setActiveTab] = useState('fase1');

  // Form states for Fase 1
  const [loadIds, setLoadIds] = useState({
    cierre_base: '684833b63f68683f4f47c7c3',
    cierre_up: '6847dafd3f68683f4f47c7c1',
    cierre_dwn: '684862413f68683f4f4e7688',
    cierre_base_efecto_curva: '684fdb65fb9753150202f140',
    cierre_up_efecto_curva: '684fdd34fb9753150207669c',
    cierre_base_efecto_balance: '684fe17afb975315020bdbf8',
    cierre_up_efecto_balance: '684fe3b7fb975315021049bd',
  });

  const [isLoading, setIsLoading] = useState(false);
  const API_URL = process.env.NEXT_PUBLIC_API_URL || 'https://tu-backend-render.com';

  const handleRunFase1 = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoading(true);
    try {
        // En una app real de producción, aquí recogerías todos los campos con FormData()
        // const formData = new FormData(e.target as HTMLFormElement);
        // formData.append('cierre_base', loadIds.cierre_base); ...
        
        const res = await fetch(`${API_URL}/api/fase1`, {
            method: 'POST',
            // body: formData
        });
        
        if(res.ok) {
            alert('¡Llamada exitosa al backend de Render (Fase 1)!');
        } else {
            alert('Fallo en la llamada al backend. Revisa los logs de Render.');
        }
    } catch (err) {
        alert(`Error conectando con el backend: ${err}`);
    } finally {
        setIsLoading(false);
    }
  };

  const handleRunFase2 = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoading(true);
    try {
        const res = await fetch(`${API_URL}/api/fase2`, {
            method: 'POST',
            // body: formData
        });
        if(res.ok) {
            alert('¡Llamada exitosa al backend de Render (Fase 2)!');
        } else {
            alert('Fallo en la llamada al backend. Revisa logs.');
        }
    } catch (err) {
        alert(`Error conectando con el backend: ${err}`);
    } finally {
        setIsLoading(false);
    }
  };

  return (
    <div>
      <div className="tabs-container">
        <button 
          className={`tab ${activeTab === 'fase1' ? 'active' : ''}`}
          onClick={() => setActiveTab('fase1')}
        >
          Fase 1: Extracción de Datos
        </button>
        <button 
          className={`tab ${activeTab === 'fase2' ? 'active' : ''}`}
          onClick={() => setActiveTab('fase2')}
        >
          Fase 2: Generar IA y PPTX
        </button>
      </div>

      {activeTab === 'fase1' && (
        <div className="glass-panel slide-top">
          <h2 style={{ marginBottom: '16px', color: 'var(--primary)' }}>Actualizar Informes Excel</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '32px' }}>
            Esta fase conectará a AWS Athena para extraer los datos utilizando los correspondientes Load IDs y generar las plantillas para la Fase 2.
          </p>

          <form onSubmit={handleRunFase1}>
            <div className="grid-2">
              <div className="glass-panel" style={{ background: 'rgba(0,0,0,0.2)' }}>
                <h3 style={{ marginBottom: '16px' }}>Plantillas de Excel</h3>
                <div className="input-group">
                  <label className="input-label">ID de Plantilla COAP (Google Drive)</label>
                  <input type="text" className="input-field" placeholder="Google Drive ID" defaultValue="ID-AQUI" />
                </div>
                <div className="input-group">
                  <label className="input-label">Plantilla_Efecto_Balance_Curva.xlsx</label>
                  <input type="file" className="input-field" accept=".xlsx" />
                </div>
                <div className="input-group">
                  <label className="input-label">Plantilla_Datos_Medios.xlsx</label>
                  <input type="file" className="input-field" accept=".xlsx" />
                </div>
              </div>

              <div className="glass-panel" style={{ background: 'rgba(0,0,0,0.2)' }}>
                <h3 style={{ marginBottom: '16px' }}>Load IDs (Athena)</h3>
                {Object.entries(loadIds).map(([key, value]) => (
                  <div key={key} className="input-group">
                    <label className="input-label" style={{ textTransform: 'capitalize' }}>
                      {key.replace(/_/g, ' ')}
                    </label>
                    <input 
                      type="text" 
                      className="input-field" 
                      value={value}
                      onChange={(e) => setLoadIds({...loadIds, [key]: e.target.value})}
                    />
                  </div>
                ))}
              </div>
            </div>
            
            <div style={{ marginTop: '32px', textAlign: 'right' }}>
              <button type="submit" className="btn">
                ▶️ Ejecutar Fase 1
              </button>
            </div>
          </form>
        </div>
      )}

      {activeTab === 'fase2' && (
        <div className="glass-panel slide-top">
          <h2 style={{ marginBottom: '16px', color: 'var(--accent)' }}>Generación de Comentarios IA</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '32px' }}>
            En esta fase, la inteligencia artificial de Gemini procesará los datos obtenidos y creará presentaciones actualizadas con comentarios y guiones tipo podcast.
          </p>

          <form onSubmit={handleRunFase2}>
            <div className="grid-2">
              <div className="glass-panel" style={{ background: 'rgba(0,0,0,0.2)' }}>
                <h3 style={{ marginBottom: '16px' }}>Archivos Generados (Input)</h3>
                <div className="input-group">
                  <label className="input-label">Archivo ALCO del mes actual (.xlsx)</label>
                  <input type="file" className="input-field" accept=".xlsx" />
                </div>
                <div className="input-group">
                  <label className="input-label">Presentación COAP Anterior (.pptx)</label>
                  <input type="file" className="input-field" accept=".pptx" />
                </div>
                <div className="input-group">
                  <label className="input-label">Plantilla Base (.pptx)</label>
                  <input type="file" className="input-field" accept=".pptx" />
                </div>
              </div>
              
              <div className="glass-panel" style={{ background: 'rgba(0,0,0,0.2)' }}>
                <h3 style={{ marginBottom: '16px' }}>Opciones</h3>
                <div className="input-group">
                  <label className="input-label">Mes de Cierre (ej: Mayo 2025)</label>
                  <input type="text" className="input-field" defaultValue="Mayo 2025" />
                </div>
                <div className="input-group">
                  <label className="input-label">API Key de Gemini</label>
                  <input type="password" className="input-field" placeholder="AIzaSy..." />
                </div>
              </div>
            </div>
            
            <div style={{ marginTop: '32px', textAlign: 'right' }}>
              <button type="submit" className="btn" style={{ background: 'var(--accent)' }}>
                ✨ Generar IA y PPTX
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}
