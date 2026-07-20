# Dados de Exemplo - Child Ticket

## Como usar

O sistema possui dois modos de operação:

### 1. Dados Mock (Automático) ✅
Quando você acessa o sistema pela **primeira vez**, os dados mock são exibidos automaticamente como exemplo:

#### Perfil: Gestante
- **21768 - Maria Evilene Menezes Oliveira** | Puérpera | Dieta: Geral
- **32083 - Antonia de Araujo Silva** | Puérpera | Dieta: Geral HAS DM

#### Perfil: Oncológico
- **40521 - Carlos Eduardo Menezes** | Oncológico | Dieta: Branda Hipercalórica

#### Catálogo de Itens (Compartilhado)
```
CAFÉ COMPLETO
CAFÉ COMPLETO DM
MINGAU DE AVEIA
SUCO
SUCO DM
VITAMINA
VITAMINA DM
GERAL
GERAL HAS DM
BRANDA
BISCOITO
MINGAU
```

### 2. Dados Reais (Planilha)
Conforme você interage com o sistema:
- Adiciona novos pacientes → Dados são salvos na planilha automaticamente
- Modifica itens de dieta → Catálogo é atualizado na planilha
- Troca de perfil → Dados são mantidos separados por perfil

A planilha é criada automaticamente na primeira execução (PropertiesService armazena o ID).

## Estrutura de Dados Esperada

### Sheet: `ITENS_DIETA`
| Item |
|------|
| CAFÉ COMPLETO |
| MINGAU DE AVEIA |
| ... |

### Sheet: `17/07/2026` (Gestante)
| Prontuário / Paciente | Diagnóstico | Dieta | Desjejum - 6h | Colação - 9h | Almoço - 12h | Lanche - 15h | Jantar - 18h | Ceia - 21h | Observação |
|---|---|---|---|---|---|---|---|---|---|
| 21768 - Maria Evilene | Puérpera | Geral | CAFE COMPLETO + 1 OVO | SUCO + BISCOITO | GERAL COM FRANGO | VITAMINA + BISCOITO | BRANDA DE FRANGO | MINGAU | |

### Sheet: `ONCO - 17/07/2026` (Oncológico)
Mesmo formato, mas com prefixo `ONCO - ` no nome da sheet

## Para Popular a Planilha Manualmente

Se quiser pré-carregar dados sem usar a interface:

1. **Abra a planilha criada pelo sistema** (verifique nas Propriedades do Script)
2. **Crie uma sheet** com o nome da data (ex: `17/07/2026`)
3. **Adicione o cabeçalho**:
   ```
   Prontuário / Paciente, Diagnóstico, Dieta, Desjejum - 6h, Colação - 9h, Almoço - 12h, Lanche - 15h, Jantar - 18h, Ceia - 21h, Observação
   ```
4. **Adicione as linhas de pacientes**
5. **Populate `ITENS_DIETA`** com os itens de catálogo

## Próximos Passos

Quer que eu crie um script automático para popular a planilha via Apps Script? Posso adicionar uma função que preenche dados de exemplo com um clique.
