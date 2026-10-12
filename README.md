# 🦁 LION APP — Sistema de Gestão e Apuração de IRPF

> **Planilha Executiva de Coleta, Consolidação e Rastreabilidade Patrimonial**  
> *Baseada no arquivo `Ferramenta.xlsx`*

---

## 📌 1. Visão Geral do Projeto

O **Lion App** é uma ferramenta estruturada em Microsoft Excel desenvolvida para organizar, auditar e consolidar as informações necessárias para a **Declaração de Ajuste Anual do Imposto de Renda Pessoa Física (IRPF)**. O nome faz alusão direta ao **"Leão"**, símbolo histórico da Receita Federal do Brasil.

A ferramenta resolve um dos maiores gargalos do processo tributário: a dispersão de dados pessoais, informes bancários e comprovantes de rendimento. Projetada com layout em formato de aplicação (menu lateral fixo, botões de navegação e ausência de linhas de grade poluídas), ela atende tanto ao **contribuinte individual** quanto a **escritórios de contabilidade e planejadores financeiros** que realizam a triagem cadastral de clientes.

---

## 🏛️ 2. Arquitetura da Ferramenta (`Ferramenta.xlsx`)

O arquivo é composto por **3 abas temáticas interligadas**, organizando o fluxo de dados em uma sequência lógica:

```
[ Menu Lateral: LION APP ]
   ├── 1. Titular   ──► Dados Cadastrais, Endereço e Parâmetros
   ├── 2. Informes  ──► Saldos Bancários, Custódia e Comprovantes
   └── 3. Notas     ──► Receitas Mensais, Holerites e Categorização
```

---

### 📋 Aba 1: `Titular` — Cadastro e Informações Pessoais

Centraliza as informações básicas exigidas na ficha de identificação da Receita Federal.

* **Cabeçalho:** `1. DADOS DO TÍTULAR`
* **Instrução:** `preencha os dados da sua pessoa física abaixo`
* **Mapeamento de Campos e Células:**
  | Célula | Campo / Rótulo | Dado Exemplo (`Ferramenta.xlsx`) | Formatação / Máscara |
  | :---: | :--- | :--- | :--- |
  | **D6** | `NOME` | Rafael Monteiro Campos | Texto livre |
  | **D7** | `CPF` | 000.000.000-00 | Máscara numérica (`000"."000"."000"-"00`) |
  | **D8** | `NASCIMENTO` | 14/08/1992 | Formato de data (`dd/mm/aaaa`) |
  | **D9** | `TÍTULO DE ELEITOR` | 0000 0000 0000 | Formato numérico/texto |
  | **D10** | `CÔNJUGE` | Marina Alves Rocha | Texto livre |
  | **D11** | `RUA` | Rua das Acácias, 145, Jardim Primavera | Texto livre |
  | **D12** | `RUA ABREVIADA` | R. das Acácias, 145 | Texto livre |
  | **D13** | `CEP` | 00000-000 | Máscara de CEP (`00000\-000`) |
  | **D14** | `TELEFONE` | (11) 3000-0000 | Máscara de telefone fixo (`(00) 0000-0000`) |
  | **D15** | `CELULAR` | (11) 90000-0000 | Máscara de celular (`(00) 00000-0000`) |
  | **D16** | `E-MAIL` | rafael.campos@example.com | Hiperlink interativo (`mailto:`) |
  | **D17** | `HOUVE ALTERAÇÕES ANTERIORES`| SIM | Validação suspensa (`SIM, NÃO`) |
  | **D18** | `DEPENDENTE CÔNJUGE` | NÃO | Validação suspensa (`SIM, NÃO`) |
  | **D19** | `RESIDENTE NO EXTERIOR` | NÃO | Validação suspensa (`SIM, NÃO`) |

* **Validação de Dados:** As células `D17:D19` possuem lista de seleção restrita às opções `SIM` e `NÃO`, prevenindo erros de digitação.
* **Navegação:** Botão `PRÓXIMO ->` direciona o usuário diretamente para a aba `Informes`.

---

### 🏦 Aba 2: `Informes` — Saldos Bancários e Custódia Patrimonial

Consolida as posições patrimoniais em contas bancárias, poupança e investimentos em 31 de dezembro.

* **Cabeçalho:** `2. INFORMES DE RENDIMENTOS BANCÁRIOS`
* **Instrução:** `Preencha com seus dados atuais de cada banco`
* **Card de Totalização:**
  * **Célula C7:D7 (Mesclada):** Fórmula ativa `=SUM(D11,D16,D21)`
  * **Valor Consolidado:** `R$ 58.377.777,00`
  * **Formato:** Moeda contábil brasileira (`"R$" #,##0.00`)
* **Estrutura dos Blocos Bancários:**
  1. **1º Banco (C6 Bank):**
     * Banco: `336 - C6 Bank` (D10)
     * Saldo / Valor Atual: `R$ 500.000,00` (D11)
     * Anexo Comprobatório: `topazao_2025.pdf` (D12)
  2. **2º Banco (Santander Conta Corrente):**
     * Banco: `33 - Banco Santander` (D15)
     * Saldo / Valor Atual: `R$ 1.312.312,00` (D16)
     * Anexo Comprobatório: `topazao_2025.pdf` (D17)
  3. **3º Banco (Santander Investimentos):**
     * Banco: `33 - Banco Santander` (D20)
     * Saldo / Valor Atual: `R$ 56.565.465,00` (D21)
     * Anexo Comprobatório: `topazao_2025.pdf` (D22)
* **Navegação:** Botão `<- ANTERIOR` (retorna para `Titular`) e `PRÓXIMO ->` (avança para `Notas`).

---

### 💼 Aba 3: `Notas` — Receitas Mensais e Holerites

Registra mês a mês as entradas tributáveis de rendimentos do contribuinte (salários, pró-labore, notas de serviços, etc.).

* **Cabeçalho:** `3. NOTAS BANCÁRIAS OU EXTRATO DE HOLERITES`
* **Instrução:** `São todos os valores de entrada mês a mês de receita`
* **Faixa de Destaque:** `ENTRADAS` (C7:E7 mesclado)
* **Tabela Estruturada (`Tabela1` — C8:E35):**
  * Tabela nativa do Excel com estilo integrado (`TableStyleMedium13`), filtros automáticos e linhas com preenchimento zebrado.
  * **Colunas:**
    1. `DATA` (Coluna C): Data do recebimento formatada como `mês-ano` (ex.: `April-2025`).
    2. `CATEGORIA` (Coluna D): Tipo de receita com validação por lista suspensa (`HOLERITE, CNPJ, FREELANCE`).
    3. `VALOR` (Coluna E): Montante bruto recebido com formatação de moeda (`R$ 3.000,00`).
* **Linhas Reservadas:** A tabela possui 27 linhas prontas (C9 a E35) para inclusão sequencial de novos lançamentos ao longo de todo o ano-calendário.
* **Navegação:** Botão `<- ANTERIOR` (retorna para `Informes`).

---

## 🎨 3. Identidade Visual e Interface de Aplicação

A ferramenta foi construída para simular um dashboard de software financeiro executivo:

1. **Paleta de Cores:**
   * **Dourado Base / Linhas de Acento:** `#B8935A`
   * **Dourado Médio:** `#D4B483`
   * **Dourado Claro / Fundo de Destaques:** `#E6D3A3`
   * **Azul-Noite (Navy Escuro):** `#1C2431` (utilizado para títulos, faixas executivas e texto sobre dourado)
   * **Creme Suave de Entrada:** `#FAF5E8`
2. **Menu Lateral Persistente:**
   * Logotipo do Leão em relevo metálico dourado.
   * Banner institucional **LION APP** com degradê horizontal.
   * Botões de abas estilizados com estados dinâmicos (aba ativa com degradê dourado brilhante; abas inativas em azul-ardósia escuro).
   * Hiperlinks de clique direto para `#Titular!A1`, `#Informes!A1` e `#Notas!A1`.
3. **Ergonomia Visual:**
   * Linhas de grade tradicionais desativadas (`showGridLines = False`), criando aspecto limpo e profissional.
   * Configuração de página em modo Paisagem (*Landscape*) ajustada para 1 página de largura por 1 página de altura (*Fit to Page*), garantindo exportação em PDF perfeita.

---

## 🚀 4. Guia Rápido de Utilização

### Passo 1: Preenchimento Cadastral (`Titular`)
1. Acesse a aba **Titular**.
2. Preencha os campos em branco com seus dados atualizados.
3. Nas linhas 17, 18 e 19, clique na seta suspensa para selecionar `SIM` ou `NÃO`.
4. Ao concluir, clique no botão **`PRÓXIMO ->`**.

### Passo 2: Consolidação Bancária (`Informes`)
1. Com os informes de rendimentos fornecidos por seus bancos em mãos, insira o nome da instituição, o saldo existente em 31/12 e o nome do arquivo PDF comprobatório.
2. A célula **TOTAL** calculará automaticamente a soma de todas as contas cadastradas (`=SUM(D11,D16,D21)`).
3. Salve os arquivos de informe na mesma pasta da planilha com os nomes indicados na coluna `ANEXO`.
4. Clique no botão **`PRÓXIMO ->`**.

### Passo 3: Registro de Receitas (`Notas`)
1. Acesse a aba **Notas**.
2. Na coluna **DATA**, insira a data da entrada (ex.: data do pagamento do holerite).
3. Na coluna **CATEGORIA**, clique na seta da lista suspensa e selecione:
   * `HOLERITE`: para rendimentos assalariados (CLT/estatutário).
   * `CNPJ`: para faturamento de pessoa jurídica / pró-labore.
   * `FREELANCE`: para trabalhos autônomos ou sem vínculo empregatício.
4. Na coluna **VALOR**, digite o valor bruto recebido.
5. Novas linhas na tabela expandem automaticamente as formatações e regras de validação.

---

## 🛡️ 5. Boas Práticas e Prevenção de Malha Fina

* **Conferência de Centavos:** Ao lançar os valores na aba `Informes`, certifique-se de preencher o valor exato apontado no documento oficial do banco. Divergências de centavos são motivos comuns de retenção em malha fiscal.
* **Guarda dos Comprovantes:** Mantenha os arquivos listados na coluna `ANEXO` arquivados por no mínimo **5 anos** (prazo decadencial da Receita Federal).
* **Consistência de Fontes Pagadoras:** A soma dos valores da aba `Notas` deve bater com os informes anuais de rendimentos emitidos pelas respectivas empresas empregadoras ou tomadoras de serviço.

---

*Arquivo gerado e documentado com base rigorosa na estrutura oficial de `Ferramenta.xlsx`.*
