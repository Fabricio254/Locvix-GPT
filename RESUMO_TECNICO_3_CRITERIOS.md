# 🔧 RESUMO TÉCNICO — Dashboard Manutenção 3 Critérios
**Desenvolvido em:** 10 de janeiro de 2025  
**Status:** ✅ Implementação Completa — Dashboard + Backend  
**Versão:** 2.0 (3-Criteria Support)

---

## 📊 ESTRUTURA DE DADOS

### Fluxo de Dados
```
FullTrack API
    ↓ (horímetro_atual, hodômetro_atual)
Python: calcular_status_manutencao()
    ↓ (Supabase: intervalo_horas, intervalo_km, periodo_dias)
raw_manutencoes[] → JSON embarcado no HTML
    ↓
JavaScript: mkManutencao() → renderiza tabela + KPIs
```

### Supabase — Tabela: `manutencoes_equipamentos`

#### Campos Novos (Migração SQL)
| Campo | Tipo | Nullable | Descrição |
|-------|------|----------|-----------|
| `hodometro_ultima_manutencao` | NUMERIC(12,2) | SIM | Hodômetro (km) última manut. |
| `intervalo_km` | NUMERIC(12,2) | SIM | Intervalo manutenção em km |
| `periodo_dias` | INTEGER | SIM | Intervalo manutenção em dias |
| `ultima_manutencao` | DATE | SIM | Data última manutenção |

#### Campos Existentes (Mantidos)
- `equipamento` (PK) — identificador único
- `horimetro_ultima_manutencao` — horímetro última manut.
- `intervalo_horas` — intervalo manut. em horas
- `tipo_servico` — descrição do serviço
- `created_at`, `updated_at` — timestamps

---

## 🐍 BACKEND (Python)

### Função Principal: `calcular_status_manutencao()`
**Localização:** `z:\codigos\Locvix GPT\locvix.py`, linhas ~2478–2607

**Assinatura:**
```python
def calcular_status_manutencao(rec, horo_atual, hodo_atual):
    """
    Calcula status de manutenção para 3 critérios independentes.
    
    Args:
        rec (dict): Registro de manutenção com intervalo_horas, intervalo_km, periodo_dias
        horo_atual (float): Horímetro atual do equipamento
        hodo_atual (float): Hodômetro atual do equipamento
    
    Returns:
        dict: {
            "status_geral": "vencida|proxima|ok",
            "criterio_urgente": "horas|km|dias|None",
            "status_horas": {"status": "vencida|proxima|ok", "restantes": float},
            "status_km": {"status": "vencida|proxima|ok", "restantes": float},
            "status_dias": {"status": "vencida|proxima|ok", "restantes": int}
        }
    """
```

**Lógica:**
1. Para cada critério (h/km/d):
   - Calcula `próxima_manut = última_manut + intervalo`
   - Calcula `restantes = próxima_manut - atual`
   - Status:
     - Se `restantes < 0` → "vencida" (urgência = 3)
     - Se `0 ≤ restantes ≤ 20% intervalo` → "proxima" (urgência = 2)
     - Se `restantes > 20% intervalo` → "ok" (urgência = 1)

2. Seleciona critério com MAIOR urgência → `status_geral` e `criterio_urgente`

### Integração em `gerar_dashboard_html()`
**Localização:** linhas ~2960–2980

```python
# Loop de equipamentos (após fetch de manutencoes)
for i, rec in enumerate(raw_manutencoes):
    # Busca horímetro/hodômetro atual do FullTrack
    veic = next((v for v in veiculos_ft if v['cc'] == rec['equipamento']), {})
    horo_atual = veic.get('horimetro', 0)
    hodo_atual = veic.get('hodometro', 0)
    
    # Calcula status multi-critério
    status_result = calcular_status_manutencao(rec, horo_atual, hodo_atual)
    
    # Adiciona resultado ao registro
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

### Função: `buscar_manutencoes()`
**Localização:** linhas ~2375–2405

**Mudanças:**
- Query Supabase retorna 3 novos campos
- Fallback para `None` se não configurados

```python
SELECT equipamento, horimetro_ultima_manutencao, intervalo_horas,
       hodometro_ultima_manutencao, intervalo_km, periodo_dias,
       tipo_servico, ultima_manutencao, updated_at
FROM manutencoes_equipamentos
WHERE id_loja = ...
```

### Função: `salvar_manutencao()`
**Localização:** linhas ~2407–2476

**Assinatura Nova:**
```python
def salvar_manutencao(equipamento, horimetro_ultima=None, intervalo_horas=None,
                      hodometro_ultima=None, intervalo_km=None, periodo_dias=None,
                      tipo_servico=None, id_loja=LOJA_DEFAULT):
    """Salva registro de manutenção com até 3 critérios."""
    payload = {
        'equipamento': equipamento,
        'horimetro_ultima_manutencao': horimetro_ultima,
        'intervalo_horas': intervalo_horas,
        'hodometro_ultima_manutencao': hodometro_ultima,
        'intervalo_km': intervalo_km,
        'periodo_dias': periodo_dias,
        'tipo_servico': tipo_servico,
        'ultima_manutencao': datetime.date.today().isoformat(),
        'id_loja': id_loja
    }
    # Supabase PATCH/POST...
```

---

## 🎨 FRONTEND (JavaScript)

### Função: `mkManutencao()`
**Localização:** linhas ~4833–4890

**Renderização de Tabela:**
```javascript
data.forEach(r => {
    // 1. Determina badge de status geral
    const badge = r.status_geral === 'vencida' ? '🔴 VENCIDA'
                : r.status_geral === 'proxima' ? '⚠️ PRÓXIMA'
                : '✅ EM DIA';
    
    // 2. Renderiza status de cada critério (se configurado)
    let statusHtml = '';
    if (r.status_horas?.status) {
        statusHtml += `<div ...>${icon} Horas: ${restantes}</div>`;
    }
    if (r.status_km?.status) {
        statusHtml += `<div ...>${icon} KM: ${restantes}</div>`;
    }
    if (r.status_dias?.status) {
        statusHtml += `<div ...>${icon} Dias: ${restantes}</div>`;
    }
    
    // 3. Cria linha da tabela
    tr.innerHTML = `
        <td>${r.cc}</td>
        <td>${r.placa}</td>
        <td>${badge}</td>
        <td>${fmtH(r.horimetro_atual)}</td>
        <td>${fmtKm(r.hodometro_atual)}</td>
        <td>${statusHtml}</td>
        <td>${r.tipo_servico} (via ${r.criterio_urgente})</td>
    `;
});
```

### Função: `salvarManutencao()`
**Localização:** linhas ~4909–4962

**Coleta 3 critérios do formulário:**
```javascript
const horoUlt = document.getElementById('mFormHoroUlt').value;
const hodoUlt = document.getElementById('mFormHodoUlt').value;
const periodo = document.getElementById('mFormPeriodoDias').value;

// Envia ao Supabase (via REST)
const payload = {
    equipamento: equip,
    horimetro_ultima_manutencao: horoUlt ? parseFloat(horoUlt) : null,
    intervalo_horas: intHoras || null,
    hodometro_ultima_manutencao: hodoUlt ? parseFloat(hodoUlt) : null,
    intervalo_km: intKm || null,
    periodo_dias: periodo ? parseInt(periodo) : null,
    tipo_servico: serv || null,
    ultima_manutencao: today,
    updated_at: now
};

fetch(sbUrl + '/rest/v1/manutencoes_equipamentos', {
    method: 'POST',
    headers: hdrs,
    body: JSON.stringify(payload)
});
```

### Função: `deletarManutencao()`
**Localização:** linhas ~4965–5003

- DELETE via WHERE equipamento=...
- Remove da lista local `MANUTENCAO[]`
- Limpa todos os campos do formulário

---

## 📋 FORMULÁRIO HTML

**Localização:** linhas ~3770–3820

### Campos de Entrada

| ID | Label | Tipo | Padrão | Obrigatório |
|----|----|------|--------|-------------|
| `mFormEquip` | Equipamento | select | — | SIM |
| `mFormHoroUlt` | Hor. Últ. Manut. (h) | number | — | NÃO* |
| `mFormIntHoras` | Inter. (h) | number | 600 | NÃO |
| `mFormHodoUlt` | Hod. Últ. (km) | number | — | NÃO* |
| `mFormIntKm` | Interval. KM | number | 5000 | NÃO |
| `mFormPeriodoDias` | Per. (dias) | number | 90 | NÃO* |
| `mFormServico` | Tipo de Serviço | text | — | NÃO |

*Pelo menos UM dos 3 deve ser preenchido

### KPIs
- `kManutVencidas` — count(status_geral='vencida')
- `kManutProximas` — count(status_geral='proxima')
- `kManutOk` — count(status_geral='ok')

### Tabela
| Coluna | Fonte | Largura |
|--------|------|---------|
| Equipamento | `r.cc` | 18% |
| Placa | `r.placa` | 7% |
| Status Geral | badge | 9% |
| Hor. Atual | `r.horimetro_atual` | 9% |
| Hod. Atual | `r.hodometro_atual` | 9% |
| Status por Critério | `r.status_horas|_km|_dias` | 35% |
| Tipo de Serviço | `r.tipo_servico` | 13% |

---

## 📝 ALTERAÇÕES DE ARQUIVO

### `z:\codigos\Locvix GPT\locvix.py`
```diff
+ Linha ~2375-2405: buscar_manutencoes() — adiciona 3 novos campos
+ Linha ~2407-2476: salvar_manutencao() — parâmetros 9 (era 4)
+ Linha ~2478-2607: calcular_status_manutencao() — FUNÇÃO NOVA
+ Linha ~2960-2980: gerar_dashboard_html() — chama calcular_status_manutencao()
+ Linha ~3770-3820: HTML formulário — 5 novos inputs
+ Linha ~3835-3855: HTML tabela — novo titulo e estrutura
+ Linha ~4833-4890: mkManutencao() — renderiza 3 critérios
+ Linha ~4909-4962: salvarManutencao() — salva 3 critérios
+ Linha ~4965-5003: deletarManutencao() — atualiza formulário
+ Linha ~6021-6036: event listener mudança de equipamento — popula 5 novos campos
```

### Arquivos Novos
```
z:\codigos\Locvix GPT\MIGRACAO_SUPABASE_3_CRITERIOS.sql
z:\codigos\Locvix GPT\README_MANUTENCAO_3_CRITERIOS.md
z:\codigos\Locvix GPT\RESUMO_TECNICO_3_CRITERIOS.md (este arquivo)
```

---

## 🔌 INTEGRAÇÃO COM `alertas_manutencao.py`

**Status Atual:** ❌ NÃO ATUALIZADO AINDA

**Mudanças Necessárias:**
1. Importar `calcular_status_manutencao()` do locvix.py
2. Para cada equipamento, chamar: `status = calcular_status_manutencao(rec, horo, hodo)`
3. Enviar alerta se: `status['status_geral'] == 'vencida'`
4. Email deve indicar: `via {criterio_urgente}` (horas/km/dias)

**Localização:** `z:\codigos\Locvix GPT\alertas_manutencao.py`

---

## ✅ CHECKLIST DE IMPLEMENTAÇÃO

### Backend
- [x] Função `calcular_status_manutencao()` criada
- [x] Parâmetros Supabase expandidos
- [x] Loop de processamento atualizado
- [x] Fallback para `None` se critério não configurado

### Frontend — HTML
- [x] Formulário: 7 campos (equipamento, hor_ult, int_h, hodo_ult, int_km, per_d, serv)
- [x] KPIs: 3 cards (vencidas, próximas, ok)
- [x] Tabela: 7 colunas com status multi-critério

### Frontend — JavaScript
- [x] `mkManutencao()` — renderiza todos os 3 critérios
- [x] `salvarManutencao()` — coleta e envia 3 critérios
- [x] `deletarManutencao()` — remove registros
- [x] Event listener — popula 5 campos ao mudar equipamento

### Supabase
- [x] Arquivo migração SQL criado: `MIGRACAO_SUPABASE_3_CRITERIOS.sql`
- [ ] ❌ Migração precisa ser EXECUTADA pelo usuário

### Documentação
- [x] README com exemplos de uso
- [x] Resumo técnico (este arquivo)

### Alerts (TODO)
- [ ] Atualizar `alertas_manutencao.py`
- [ ] Integrar `calcular_status_manutencao()` 
- [ ] Enviar 3 tipos de alerta (h/km/d)

---

## 🚀 PRÓXIMOS PASSOS

1. **Executar Migração SQL** — Abrir Supabase > SQL Editor > colar arquivo > Run
2. **Testar Dashboard** — Selecionar equipamento > preencher 3 critérios > salvar
3. **Validar Tabela** — Verificar que 3 status cards aparecem
4. **Atualizar Alertas** — Replicar lógica `calcular_status_manutencao()` em `alertas_manutencao.py`

---

## 📞 DEBUGGING

**Erro ao salvar:**
- F12 > Console > verifique se há erro de fetch
- Verifique: `_SB_URL` e `_SB_ANON` configurados

**Status não atualiza:**
- Verifique que pelo menos 1 critério está preenchido
- Reload da página (F5) para carregar dados novos

**Tabela vazia:**
- Abra Supabase > `manutencoes_equipamentos` > verifique registros
- Cheque permissões RLS (row-level security)

---

**Fim do Documento**  
Desenvolvido em Python/JavaScript para Locvix Dashboard v2.0
