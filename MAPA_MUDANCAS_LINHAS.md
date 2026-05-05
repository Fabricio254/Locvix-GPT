# 📝 MAPA DE MUDANÇAS — Arquivo: locvix.py

## 🔍 Onde Procurar Cada Mudança

---

## 🐍 PYTHON — Backend

### 1️⃣ Função `buscar_manutencoes()`
**Localização:** ~Linhas 2375–2405  
**Mudança:** Query Supabase retorna 3 novos campos  
**Código-chave:**
```python
SELECT equipamento, horimetro_ultima_manutencao, intervalo_horas,
       hodometro_ultima_manutencao, intervalo_km, periodo_dias,  # ← NOVOS
       tipo_servico, ultima_manutencao, updated_at
FROM manutencoes_equipamentos
```

---

### 2️⃣ Função `salvar_manutencao()`
**Localização:** ~Linhas 2407–2476  
**Mudança:** Assinatura expandida de 4 para 9 parâmetros  
**Novo:** Suporta hodometro, intervalo_km, periodo_dias  
```python
def salvar_manutencao(equipamento,
    horimetro_ultima=None,        # existente
    intervalo_horas=None,         # existente
    hodometro_ultima=None,        # NOVO
    intervalo_km=None,            # NOVO
    periodo_dias=None,            # NOVO
    tipo_servico=None,            # existente
    id_loja=LOJA_DEFAULT):        # existente
```

---

### 3️⃣ Função `calcular_status_manutencao()` ⭐
**Localização:** ~Linhas 2478–2607  
**Status:** ✨ FUNÇÃO COMPLETAMENTE NOVA  
**Função:** Calcula status para 3 critérios independentes  
**Retorna:**
```python
{
    "status_geral": "vencida|proxima|ok",
    "criterio_urgente": "horas|km|dias|None",
    "status_horas": {"status": "...", "restantes": float},
    "status_km": {"status": "...", "restantes": float},
    "status_dias": {"status": "...", "restantes": int}
}
```

---

### 4️⃣ Loop em `gerar_dashboard_html()`
**Localização:** ~Linhas 2960–2980  
**Mudança:** Chama `calcular_status_manutencao()` para cada equipamento  
**Código-chave:**
```python
for i, rec in enumerate(raw_manutencoes):
    veic = next((v for v in veiculos_ft if v['cc'] == rec['equipamento']), {})
    horo_atual = veic.get('horimetro', 0)
    hodo_atual = veic.get('hodometro', 0)
    
    status_result = calcular_status_manutencao(rec, horo_atual, hodo_atual)
    rec.update({
        'horimetro_atual': horo_atual,
        'hodometro_atual': hodo_atual,
        'status_geral': status_result['status_geral'],
        'criterio_urgente': status_result['criterio_urgente'],
        'status_horas': status_result['status_horas'],
        'status_km': status_result['status_km'],
        'status_dias': status_result['status_dias']
    })
```

---

## 🎨 HTML — Formulário e Tabela

### 5️⃣ Formulário Manutenção
**Localização:** ~Linhas 3768–3825  
**Mudança:** Expandido de 4 campos para 7  

#### Campos Originais (Mantidos):
- `mFormEquip` — Equipamento (select)
- `mFormHoroUlt` — Horímetro Última Manut.
- `mFormIntHoras` — Intervalo (h)
- `mFormServico` — Tipo de Serviço

#### Campos Novos:
- `mFormHodoUlt` — Hodômetro Última Manut. (km)
- `mFormIntKm` — Intervalo (km)
- `mFormPeriodoDias` — Período (dias)

---

### 6️⃣ Tabela HTML
**Localização:** ~Linhas 3835–3855  
**Mudança:** Novo título + estrutura de 7 colunas  

**Antigo:**
```html
<h3>📋 Status de Manutenção por Equipamento — Horímetro (FullTrack v24)</h3>
<!-- 8 colunas -->
```

**Novo:**
```html
<h3>📋 Status de Manutenção — 3 Critérios: Horímetro (h) + Hodômetro (km) + Período (dias)</h3>
<!-- 7 colunas com colgroup ajustado -->
```

**Colunas:**
1. Equipamento (18%)
2. Placa (7%)
3. Status Geral (9%)
4. Hor. Atual (9%)
5. Hod. Atual (9%)
6. Status por Critério (35%)
7. Tipo de Serviço (13%)

---

## 🔧 JavaScript — Renderização

### 7️⃣ Função `mkManutencao()`
**Localização:** ~Linhas 4833–4890  
**Mudança:** Renderiza 3 critérios com status cards  

**Principais alterações:**
- Usa `r.status_geral` ao invés de `r.status`
- Renderiza 3 cards (se configurados):
  - `r.status_horas` → 🔴/⚠️/✅ Horas: [restantes]
  - `r.status_km` → 🔴/⚠️/✅ KM: [restantes]
  - `r.status_dias` → 🔴/⚠️/✅ Dias: [restantes]
- Adiciona coluna com `statusHtml` (cards aninhados)

---

### 8️⃣ Função `salvarManutencao()`
**Localização:** ~Linhas 4909–4962  
**Mudança:** Coleta 5 novos inputs + valida critérios  

**Coleta:**
```javascript
const horoUlt = document.getElementById('mFormHoroUlt').value;
const hodoUlt = document.getElementById('mFormHodoUlt').value;
const intKm = document.getElementById('mFormIntKm').value;
const periodo = document.getElementById('mFormPeriodoDias').value;
const intHoras = document.getElementById('mFormIntHoras').value;
```

**Validação:** Exige pelo menos 1 de (horoUlt, hodoUlt, periodo)

**Payload ao Supabase (9 campos):**
```javascript
{
    equipamento: equip,
    horimetro_ultima_manutencao: horoUlt ? parseFloat(horoUlt) : null,
    intervalo_horas: intHoras || null,
    hodometro_ultima_manutencao: hodoUlt ? parseFloat(hodoUlt) : null,
    intervalo_km: intKm || null,
    periodo_dias: periodo ? parseInt(periodo) : null,
    tipo_servico: serv || null,
    ultima_manutencao: new Date().toISOString().slice(0,10),
    updated_at: new Date().toISOString()
}
```

---

### 9️⃣ Função `deletarManutencao()`
**Localização:** ~Linhas 4965–5003  
**Mudança:** Remove referência a campo inexistente + limpa 5 novos campos  

**Removido:**
```javascript
document.getElementById('mFormData').value = new Date()... ❌ (esse ID não existe)
```

**Adicionado:**
```javascript
document.getElementById('mFormHodoUlt').value = '';
document.getElementById('mFormIntKm').value = '5000';
document.getElementById('mFormPeriodoDias').value = '90';
```

---

### 🔟 Event Listener — DOMContentLoaded
**Localização:** ~Linhas 6021–6036  
**Mudança:** Popula 5 novos campos ao mudar equipamento  

**Antigo:**
```javascript
mSel.addEventListener('change', () => {
    const rec = MANUTENCAO.find(r => r.cc === mSel.value);
    horoEl.value = (rec && rec.horimetro_ultima != null) ? rec.horimetro_ultima : '';
    intEl.value = (rec && rec.intervalo_horas) ? rec.intervalo_horas : '600';
    srv.value = (rec && rec.tipo_servico) ? rec.tipo_servico : '';
});
```

**Novo:**
```javascript
mSel.addEventListener('change', () => {
    const rec = MANUTENCAO.find(r => r.cc === mSel.value);
    horoEl.value = (rec && rec.horimetro_ultima_manutencao != null) 
        ? rec.horimetro_ultima_manutencao : '';
    intEl.value = (rec && rec.intervalo_horas) ? rec.intervalo_horas : '600';
    
    hodoEl.value = (rec && rec.hodometro_ultima_manutencao != null)
        ? rec.hodometro_ultima_manutencao : '';  // NOVO
    intKmEl.value = (rec && rec.intervalo_km) ? rec.intervalo_km : '5000';  // NOVO
    periEl.value = (rec && rec.periodo_dias) ? rec.periodo_dias : '90';  // NOVO
    
    srv.value = (rec && rec.tipo_servico) ? rec.tipo_servico : '';
});
```

---

## 📁 ARQUIVOS ADICIONADOS (Não estão em locvix.py)

```
z:\codigos\Locvix GPT\
├── MIGRACAO_SUPABASE_3_CRITERIOS.sql          ← SQL para Supabase
├── README_MANUTENCAO_3_CRITERIOS.md           ← Guia do usuário
├── RESUMO_TECNICO_3_CRITERIOS.md              ← Doc técnica
└── IMPLEMENTACAO_CONCLUIDA.txt                ← Este sumário
```

---

## 🎯 CHECKLIST DE REVISÃO

- [ ] Linhas 2375–2405: `buscar_manutencoes()` retorna 3 novos campos
- [ ] Linhas 2407–2476: `salvar_manutencao()` com 9 parâmetros
- [ ] Linhas 2478–2607: `calcular_status_manutencao()` criada ✨
- [ ] Linhas 2960–2980: Loop em `gerar_dashboard_html()` atualizado
- [ ] Linhas 3768–3825: Formulário com 7 campos
- [ ] Linhas 3835–3855: Tabela HTML com novo título
- [ ] Linhas 4833–4890: `mkManutencao()` renderiza 3 critérios
- [ ] Linhas 4909–4962: `salvarManutencao()` coleta 3 critérios
- [ ] Linhas 4965–5003: `deletarManutencao()` corrigida
- [ ] Linhas 6021–6036: Event listener popula 5 campos

---

## 🚀 TESTE RÁPIDO

**Para validar que tudo funciona:**

1. Abrir arquivo locvix.py em editor
2. Procurar por: `calcular_status_manutencao`
   - Deve encontrar em ~linha 2478 (função nova)
3. Procurar por: `mFormHodoUlt`
   - Deve encontrar 4 vezes (HTML + 3x JavaScript)
4. Procurar por: `status_geral`
   - Deve encontrar em Python, JavaScript e JSON
5. Procurar por: `criterio_urgente`
   - Deve encontrar em Python, JavaScript

**Se todos os 5 passarem → ✅ Implementação está correta!**

---

**Fim do Mapa de Mudanças**
