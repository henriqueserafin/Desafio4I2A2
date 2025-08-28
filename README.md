
# Automação da Compra de VR/VA

Automatiza o processo mensal de compra de VR (Vale Refeição), consolidando bases de RH, aplicando as regras de negócio e gerando a planilha final exatamente no layout do arquivo de referência `VRMENSAL05_2025.xlsx`.

## Principais Recursos
- Consolidação de múltiplas bases (Ativos, Férias, Desligados, Adm., Dias Úteis, Sindicato x Valor, Afastamentos, Estágio, Aprendiz, Exterior)
- Exclusões automáticas: estagiários, aprendizes, afastados, exterior e diretores
- Cálculo de dias úteis por colaborador considerando férias e desligamento
- Valor diário do VR por estado com fallback
- Rateio automático: 80% empresa / 20% colaborador
- Geração do Excel final sem colunas "Unnamed" e com cabeçalhos e ordem idênticos ao modelo

## Estrutura de Pastas e Arquivos
Coloque os arquivos `.xlsx` na raiz do projeto ou nas pastas `Dados/` ou `Uploads/`.

Arquivos esperados:
- `ATIVOS.xlsx`
- `FERIAS.xlsx`
- `DESLIGADOS.xlsx`
- `ADMISSOABRIL.xlsx`
- `Basesindicatoxvalor.xlsx`
- `Basediasuteis.xlsx`
- `AFASTAMENTOS.xlsx`
- `ESTAGIO.xlsx`
- `APRENDIZ.xlsx`
- `EXTERIOR.xlsx`
- `VRMENSAL05_2025.xlsx` (apenas como referência de layout)

Observações:
- Qualquer coluna contendo "matric" ou "cadastro" é padronizada para `MATRICULA`.
- Qualquer coluna contendo "sind" é padronizada para `Sindicato`.

## Regras de Negócio Implementadas
- Exclusões (por matrícula): Estagiário, Aprendiz, Afastado, Exterior, Diretor (cargo contém "DIRETOR")
- Dias úteis por colaborador: conforme `Basediasuteis.xlsx` com matching por substring do nome do sindicato
- Férias: desconto dos dias informados em `FERIAS.xlsx`
- Desligamento:
  - Se `COMUNICADO DE DESLIGAMENTO` = "OK" e data ≤ dia 15 ⇒ 0 dias (sem benefício)
  - Se data > dia 15 ⇒ dias proporcionais (proporção de 30 dias)
- Valor diário VR por estado (de `Basesindicatoxvalor.xlsx`):
  - São Paulo: 37,50 (padrão se não houver na planilha)
  - Rio de Janeiro, Rio Grande do Sul, Paraná: 35,00 (padrão se não houver na planilha)
  - Fallback geral: 35,00
- Rateio: 80% empresa / 20% colaborador

##  Como Executar
### Pré‑requisitos
- Python 3.9+
- Bibliotecas: `pandas`, `numpy`, `openpyxl`, `jupyter`

Instale as dependências:
```bash
pip install -U pandas numpy openpyxl jupyter
```

### Passos
1. Garanta que os arquivos `.xlsx` estejam na raiz, `Dados/` ou `Uploads/`.
2. Abra e rode o notebook: `VR_VA_Automacao_Modelo.ipynb`.
3. Execute as células em sequência. Ao final será gerado um arquivo Excel no padrão do modelo.

## Saída Gerada
- Arquivo: `VR_FINAL_YYYYMMDD_HHMMSS.xlsx`
- Aba: `VR Mensal`
- Colunas (ordem exata):
  - Matricula
  - Admissão
  - Sindicato do Colaborador
  - Competência
  - Dias
  - VALOR DIÁRIO VR
  - TOTAL
  - Custo empresa
  - Desconto profissional
  - OBS GERAL

## ⚙️ Parametrizações
- Competência padrão: `2025-05-01`.
  - No notebook, procure a variável `COMPETENCIA` e ajuste para o mês desejado.
- Nome da aba de saída: altere `sheet_name='VR Mensal'` no `ExcelWriter` se necessário.
- Matching leniente: sindicato (dias úteis) e estado (valor VR) são buscados por substring.

## Validações e Conferências
- O notebook imprime resumos e amostras da base final.
- Verifique:
  - Ausência de colunas `Unnamed` no Excel final
  - Nomes e ordem das colunas idênticos ao modelo
  - Totais financeiros (empresa, colaboradores, total VR)

## Solução de Problemas
- Colunas "Unnamed" no Excel:
  - O notebook salva com `index=False`. Se ainda aparecer, confirme que está abrindo o arquivo certo (sem cache) e que o viewer não adiciona cabeçalho próprio.
- Datas incorretas:
  - Garanta que colunas de data nas planilhas origem estejam como data ou número serial do Excel.
- Mapeamentos não aplicados:
  - Revise se o nome do sindicato/estado contém a substring esperada pela base de mapeamento.