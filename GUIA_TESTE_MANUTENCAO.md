# 🧪 GUIA DE TESTE — Manutenção 3 Critérios

Siga este guia para validar que a implementação está funcionando corretamente.

---

## ✅ TESTE 1: Migração Supabase

### Pré-requisito: Acesso ao Supabase

1. Abra: https://app.supabase.com
2. Login com suas credenciais
3. Selecione seu projeto Locvix

### Passo 1: Executar Script SQL

1. Vá para **SQL Editor** (menu esquerdo)
2. Clique em **New Query**
3. Abra arquivo: `z:\codigos\Locvix GPT\MIGRACAO_SUPABASE_3_CRITERIOS.sql`
4. Copie TODO o conteúdo (Ctrl+A, Ctrl+C)
5. Cole na aba de SQL do Supabase (Ctrl+V)
6. Clique em **Run** (botão azul)
7. Aguarde resultado: ✅ "Success" (sem erros)

### Validação

```
Esperado no painel de execução:
✅ execution completed successfully

Significa: 4 colunas adicionadas à tabela
```

---

## ✅ TESTE 2: Verificar Estrutura Supabase

### Visualizar Colunas Novas

1. No Supabase, vá para **Table Editor** (esquerda)
2. Selecione tabela: `manutencoes_equipamentos`
3. Verifique se as colunas abaixo aparecem:
   - ✅ `hodometro_ultima_manutencao` (numeric)
   - ✅ `intervalo_km` (numeric)
   - ✅ `periodo_dias` (integer)
   - ✅ `ultima_manutencao` (date)

**Se não aparecerem → Migração não foi executada!**

---

## ✅ TESTE 3: Testar Dashboard — Formulário

### Abrir Dashboard

1. Abra o dashboard Locvix em seu navegador
2. Vá para seção: **🛠 Manutenção Preventiva — Equipamentos**
3. Procure pelo formulário: **🔧 Registrar / Atualizar Manutenção**

### Validar Campos

Verificar se os 7 campos aparecem:

```
☐ Equipamento (dropdown com FullTrack)
☐ ⏱ Hor. Últ. Manut. (h) — input number
☐ 🔁 Inter. (h) — input number, padrão 600
☐ 🛣 Hod. Últ. (km) — input number  ← NOVO
☐ Interval. KM — input number, padrão 5000  ← NOVO
☐ 📅 Per. (dias) — input number, padrão 90  ← NOVO
☐ Tipo de Serviço — input text
```

**Se algum campo não aparecer → HTML não foi atualizado!**

### Testar Seleção de Equipamento

1. Clique na dropdown "Equipamento"
2. Selecione um equipamento (ex: ESCAV-001)
3. Observe se os campos abaixo popul automaticamente:
   - Horímetro Última Manut.
   - Intervalo (h)
   - Hodômetro Última Manut. ← NOVO
   - Intervalo KM ← NOVO
   - Período (dias) ← NOVO
   - Tipo de Serviço

**Se não popular → Event listener não foi atualizado!**

---

## ✅ TESTE 4: Salvar Registro com 1 Critério

### Preenchimento

1. Na dropdown, selecione um equipamento
2. Preencha **APENAS Horímetro**:
   - Hor. Últ. Manut.: 1000
   - Inter. (h): 600
   - Deixar outros campos em branco
3. Clique: **💾 Salvar**

### Validação

- Mensagem deve aparecer: ✅ "EQUIPAMENTO — H: 1000h registrado."
- Tabela deve atualizar com novo equipamento/horímetro
- Campo Hodômetro deve permanecer vazio (NULL no Supabase)

---

## ✅ TESTE 5: Salvar Registro com 3 Critérios

### Preenchimento

1. Selecione um novo equipamento
2. Preencha **TODOS os 3 critérios**:
   - Hor. Últ. Manut.: 2500
   - Inter. (h): 400
   - Hod. Últ. (km): 75000
   - Interval. KM: 5000
   - Per. (dias): 90
   - Tipo de Serviço: Manutenção completa
3. Clique: **💾 Salvar**

### Validação

- Mensagem: ✅ "EQUIPAMENTO — H: 2500h + KM: 75000 + Per: 90d registrado."
- Tabela mostra 3 linhas de status:
  ```
  ✅ Horas: 150h (próxima em 2900h)
  ✅ KM: 5000km (próxima em 80000km)
  ✅ Dias: 60d (próxima em 90d)
  ```

---

## ✅ TESTE 6: Validação de Status

### Cenário 1: Horímetro Vencido

1. Preencha:
   - Hor. Últ. Manut.: 1000
   - Inter. (h): **200** (intervalo pequeno)
   - Hod. Últ.: deixar vazio
   - Per. (dias): deixar vazio
2. Salvar

### Validação Esperada

- Status Geral: 🔴 **VENCIDA**
- Card de Horas deve mostrar:
  ```
  🔴 Horas: -800h (horímetro atual 1200 > 1200 de próxima manut)
  ```

### Cenário 2: Próxima Manutenção

1. Preencha:
   - Hor. Últ. Manut.: 1000
   - Inter. (h): 600
   - Hod. Últ. (km): 50000
   - Interval. KM: 5000
   - Per. (dias): 90
2. Salvar

### Validação Esperada

- Se horímetro atual = 1150:
  - Status Geral: ⚠️ **PRÓXIMA** (≤20% de 600h = ≤120h restantes)
  - Card de Horas: ⚠️ Horas: 450h

- Se hodômetro atual = 54000:
  - Status Geral: ✅ **EM DIA** (4000km > 20% de 5000km)
  - Card de KM: ✅ KM: 1000km

---

## ✅ TESTE 7: Deletar Registro

### Procedure

1. Na dropdown, selecione um equipamento que acabou de salvar
2. Clique: **🗑 Excluir**
3. Confirme no pop-up: Sim

### Validação

- Mensagem: ✅ "Registro de EQUIPAMENTO excluído."
- Formulário limpa todos os campos:
  ```
  Hor. Últ.: [vazio]
  Inter. (h): [vazio]
  Hod. Últ.: [vazio]
  Interval. KM: [vazio]
  Per. (dias): [vazio]
  Tipo de Serviço: [vazio]
  ```
- Tabela remove a linha do equipamento

---

## ✅ TESTE 8: KPIs (Contadores)

### Validação

Acima da tabela, 3 cards devem exibir:

```
🔴 Vencidas: [número]
⚠️ Próximas (≤20%): [número]
✅ Em Dia: [número]
```

- Vencidas: count(status_geral='vencida')
- Próximas: count(status_geral='proxima')
- Em Dia: count(status_geral='ok')

**Exemplo:** Se salvar 3 equipamentos (1 vencido, 1 próximo, 1 ok):
```
🔴 Vencidas: 1
⚠️ Próximas: 1
✅ Em Dia: 1
```

---

## ✅ TESTE 9: Tabela com 3 Critérios

### Verificar Renderização

A tabela deve ter 7 colunas:

| Coluna | Conteúdo Exemplo |
|--------|------------------|
| Equipamento | ESCAV-001 |
| Placa | LJ-02 |
| Status Geral | 🔴 VENCIDA (ou ⚠️ PRÓXIMA, ✅ EM DIA) |
| Hor. Atual | 1200 h |
| Hod. Atual | 75000 km |
| Status por Critério | 🔴 Horas: -100h<br/>✅ KM: 3000km<br/>✅ Dias: 65d |
| Tipo de Serviço | Troca óleo (fonte: api · via horas) |

### Validação

- Cada status card tem cor:
  - 🔴 vermelho (#dc2626)
  - ⚠️ amarelo (#d97706)
  - ✅ verde (#059669)
- Texto de "fonte: api · via [critério]" aparece

---

## ✅ TESTE 10: Teste de Erro — Campo Obrigatório

### Validação de Entrada

1. Selecione um equipamento
2. **Deixe todos os campos em branco** (ou apenas tipo de serviço)
3. Clique: **💾 Salvar**

### Validação Esperada

- Mensagem de erro: ❌ "Informe pelo menos uma métrica: horímetro, hodômetro ou período."
- Cor: vermelho (#dc2626)

---

## ✅ TESTE 11: Verificar Dados no Supabase

### Conferir Salvamento

1. Abra Supabase > Table Editor
2. Selecione: `manutencoes_equipamentos`
3. Procure pelo equipamento que salvou no Teste 5
4. Verifique colunas:
   - `equipamento` = EQUIPAMENTO_SELECIONADO
   - `horimetro_ultima_manutencao` = 2500
   - `intervalo_horas` = 400
   - `hodometro_ultima_manutencao` = 75000
   - `intervalo_km` = 5000
   - `periodo_dias` = 90
   - `tipo_servico` = "Manutenção completa"
   - `ultima_manutencao` = 2025-01-10 (data de hoje)

**Se dados não aparecer → salvarManutencao() não enviou!**

---

## ⚠️ TESTE 12: Teste de Compatibilidade — Navegador

### Verificar Console

1. Abra o dashboard
2. Pressione: F12 (abrir Developer Tools)
3. Vá para aba: **Console**
4. Procure por erros (vermelho)

### Validação Esperada

- Sem erros JavaScript
- Sem warnings relacionados a MANUTENCAO ou mkManutencao

**Se houver erro → debug via F12 > Console > leia mensagem de erro**

---

## 🎯 RESUMO DO TESTE

### ✅ Se todos os 12 testes passaram:
```
Sistema 100% funcional!
Próximo passo: atualizar alertas_manutencao.py
```

### ❌ Se algum teste falhou:

| Teste | Provável Causa |
|-------|----------------|
| 1-2 | Migração SQL não executada |
| 3 | HTML não foi atualizado |
| 4-5 | JavaScript não coleta dados |
| 6 | Cálculo de status incorreto |
| 7 | deletarManutencao() não funciona |
| 8 | mkManutencao() não conta |
| 9 | Renderização HTML incorreta |
| 10 | Validação não está implementada |
| 11 | salvarManutencao() não envia ao Supabase |
| 12 | Erro de importação ou sintaxe |

---

## 📞 Suporte

Se algum teste não passar:
1. Verifique o arquivo de changelog: `MAPA_MUDANCAS_LINHAS.md`
2. Procure a linha indicada em `locvix.py`
3. Verifique se a mudança foi aplicada corretamente
4. Se persistir: revise arquivo `RESUMO_TECNICO_3_CRITERIOS.md` para entender a lógica

---

**Fim do Guia de Teste**
