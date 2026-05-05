# 🛠 Manutenção Preventiva — 3 Critérios (Horímetro + Hodômetro + Período)

## 📋 Resumo das Mudanças

O módulo de manutenção preventiva foi **completamente refatorado** para suportar **3 critérios independentes**:

| Critério | Campo | Exemplo | Aplicação |
|----------|-------|---------|-----------|
| ⏱ **Horímetro** | Horas trabalhadas | 1.200 h | Máquinas com sensor de horas |
| 🛣 **Hodômetro** | Quilometragem | 50.000 km | Veículos (ônibus, caminhões) |
| 📅 **Período** | Dias calendário | 90 dias | Todos (preventivo por tempo) |

Cada critério é **independente e opcional**. O sistema calcula o status mais urgente automaticamente.

---

## ⚙️ Instalação / Atualização

### Passo 1: Aplicar Migração no Supabase (PRIMEIRO!)

1. Abra seu projeto no Supabase: https://app.supabase.com
2. Vá para **SQL Editor** (lado esquerdo)
3. Clique em **New Query**
4. Abra o arquivo: `MIGRACAO_SUPABASE_3_CRITERIOS.sql`
5. Copie TODO o conteúdo
6. Cole na aba de SQL do Supabase
7. Clique em **Run**
8. Espere a mensagem ✅ "Success"

**Resultado esperado:** 4 novas colunas adicionadas à tabela `manutencoes_equipamentos`:
- `hodometro_ultima_manutencao` (número, ex: 50000.50)
- `intervalo_km` (número, ex: 5000)
- `periodo_dias` (inteiro, ex: 90)
- `ultima_manutencao` (data, ex: 2025-01-10)

---

### Passo 2: Usar o Dashboard

O dashboard já foi atualizado com:

✅ **Formulário de entrada** com 5 campos:
- Equipamento (dropdown com FullTrack)
- Horímetro na última manutenção (h)
- Intervalo (h) — padrão 600h
- Hodômetro na última manutenção (km)
- Intervalo KM — padrão 5.000 km
- Período (dias) — padrão 90 dias
- Tipo de serviço

✅ **KPIs** — resumo geral:
- 🔴 Vencidas
- ⚠️ Próximas (≤20% do intervalo)
- ✅ Em Dia

✅ **Tabela de Status** — 7 colunas:
| Coluna | Descrição |
|--------|-----------|
| Equipamento | Nome/Centro de custo |
| Placa | Placa do veículo |
| Status Geral | Vencida/Próxima/OK (pelo critério mais urgente) |
| Hor. Atual | Horímetro atual (FullTrack) |
| Hod. Atual | Hodômetro atual (FullTrack, se disponível) |
| Status por Critério | Cards com status individual de cada métrica |
| Tipo de Serviço | O que foi/será feito |

---

## 🎯 Como Usar

### Exemplo 1: Máquina com Horímetro
Você tem uma **escavadeira** (FullTrack: ESCAV-001):
- Horímetro atual: **1.500 h**
- Última manutenção: **1.200 h**
- Intervalo de manutenção: **600 h**

**Campos a preencher:**
- Equipamento: `ESCAV-001`
- Hor. Últ. Manut.: `1200`
- Inter. (h): `600`
- Hod. Últ.: (deixar em branco)
- Interval. KM: (deixar em branco)
- Per. (dias): (deixar em branco)
- Tipo de Serviço: `Troca óleo e filtros`

**Resultado no status:**
- Horímetro: Restam **100 h** → Status: ✅ Em Dia
- Hodômetro: (não configurado) — ignorado
- Período: (não configurado) — ignorado
- **Status Geral: ✅ Em Dia**

### Exemplo 2: Veículo com Hodômetro + Período
Você tem um **ônibus** (FullTrack: ONIBUS-RJ-001):
- Hodômetro atual: **52.000 km**
- Última manutenção: **50.000 km**
- Intervalo KM: **5.000 km**
- Intervalo período: **90 dias**
- Última manutenção: **25 dias atrás**

**Campos a preencher:**
- Equipamento: `ONIBUS-RJ-001`
- Hor. Últ. Manut.: (deixar em branco)
- Inter. (h): (deixar em branco)
- Hod. Últ.: `50000`
- Interval. KM: `5000`
- Per. (dias): `90`
- Tipo de Serviço: `Manutenção preventiva geral`

**Resultado no status:**
- Horímetro: (não configurado) — ignorado
- Hodômetro: Restam **3.000 km** → Status: ✅ Em Dia
- Período: Restam **65 dias** → Status: ✅ Em Dia
- **Status Geral: ✅ Em Dia** (nenhum critério urgente)

### Exemplo 3: Tudo ao Mesmo Tempo
Máquina crítica com **todos os 3 critérios**:
- Equipamento: `COMPACTADOR-LJ-02`
- Hor. Últ. Manut.: `3500`
- Inter. (h): `400` (próxima a vencer!)
- Hod. Últ.: `120000`
- Interval. KM: `20000`
- Per. (dias): `180` (próximo em 30 dias)
- Tipo de Serviço: `Revisão completa com peças`

**Resultado no status:**
- Horímetro: Restam **-100 h** ➜ 🔴 **VENCIDA** (100h atrasada!)
- Hodômetro: Restam **15.000 km** → ✅ Em Dia
- Período: Restam **150 dias** → ✅ Em Dia
- **Status Geral: 🔴 VENCIDA** (critério mais urgente = horímetro)

---

## 📊 Interpretação da Tabela

Cada linha da tabela exibe:

```
Equipamento: COMPACTADOR
Placa: LJ-02
Status Geral: 🔴 VENCIDA
Hor. Atual: 3.600 h
Hod. Atual: 135.000 km

Status por Critério:
  🔴 Horas: -100 h          ← VENCIDA (100h atrasada)
  ⚠️ KM: 5.000 km           ← PRÓXIMA (dentro de 20% do intervalo)
  ✅ Dias: 150 d            ← EM DIA

Tipo de Serviço: Revisão completa com peças (fonte: api · via horas)
```

**Legenda:**
- 🔴 **Vencida**: intervalo já passou (restantes < 0)
- ⚠️ **Próxima**: faltam ≤20% do intervalo para vencer
- ✅ **Em Dia**: confortável até o próximo intervalo

---

## 🔔 Alertas Automáticos

O arquivo `alertas_manutencao.py` foi atualizado para:

1. ✅ Calcular status para **todos os 3 critérios**
2. ✅ Enviar alerta se **qualquer um** estiver vencido ou próximo
3. ✅ Indicar no email qual critério desencadeou o alerta

**Agendamento (Task Scheduler):**
```
Comando: python z:\codigos\Locvix GPT\alertas_manutencao.py
Frequência: Diariamente às 08:00
```

---

## 🛠 Troubleshooting

### P: Campos não aparecem no formulário?
**R:** Migração ainda não foi aplicada. Abra [MIGRACAO_SUPABASE_3_CRITERIOS.sql] e execute no Supabase.

### P: Salvar não está funcionando?
**R:** Verifique:
1. Supabase URL está configurado (variável `supabase_url` em locvix.py)
2. Chave anônima está configurada (variável `supabase_anon`)
3. Network do navegador (F12 > Console) não mostra erros

### P: Como voltar se der problema?
**R:** A migração só **ADICIONA** colunas (não apaga). Para reverter:
```sql
ALTER TABLE public.manutencoes_equipamentos
DROP COLUMN IF EXISTS hodometro_ultima_manutencao;
DROP COLUMN IF EXISTS intervalo_km;
DROP COLUMN IF EXISTS periodo_dias;
DROP COLUMN IF EXISTS ultima_manutencao;
```

---

## 📝 Notas Técnicas

**Backend (Python):**
- Função: `calcular_status_manutencao(rec, horo_atual, hodo_atual)`
- Retorna: dicionário com status_geral, status_horas, status_km, status_dias, criterio_urgente
- Threshold: 20h para horas, 100km para km, 5 dias para período
- Urgência: vencida (3) > próxima (2) > ok (1)

**Frontend (JavaScript):**
- Função: `mkManutencao()` — renderiza tabela com 3 critérios
- Função: `salvarManutencao()` — envia 9 campos ao Supabase
- Função: `deletarManutencao()` — remove registro

**Supabase:**
- Tabela: `manutencoes_equipamentos`
- Novas colunas: NUMERIC(12,2), INTEGER
- Valores NULL = critério não configurado para aquele equipamento

---

## 🎓 Próximas Funcionalidades (Roadmap)

- [ ] Exportar relatório mensal de manutenções por centro de custo
- [ ] Gráfico de tendência (manutenções realizadas vs planejadas)
- [ ] Integração com OrdensDeServiço (OS) para rastreabilidade
- [ ] Foto/anexo de evidência (antes/depois)
- [ ] Assinatura digital do responsável
- [ ] Sugestão automática de intervalo (baseado em histórico)

---

## 📞 Suporte

Dúvidas ou problemas? 
- Arquivo: `z:\codigos\Locvix GPT\locvix.py` (linhas 2375–2650 Python, linhas 4830–4970 JavaScript)
- Contato: fabricio.zamprognol@...
