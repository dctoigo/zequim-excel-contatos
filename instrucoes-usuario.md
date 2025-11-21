# 📋 INSTRUÇÕES DE USO - CONSOLIDADOR DE CONTATOS

## 🎯 Passo a Passo para Consolidar seus Dados

### **1️⃣ PREPARAÇÃO DAS PASTAS**

Antes de começar, organize suas pastas conforme abaixo:

#### 📁 **Pasta de Origem** (arquivos brutos)
- Crie uma pasta no seu computador para os **arquivos originais**
- Coloque **todos os arquivos Excel** que deseja processar dentro desta pasta
- ⚠️ **IMPORTANTE:** O nome da pasta **NÃO** deve conter:
  - Espaços (use `_` no lugar)
  - Caracteres especiais (acentos, ç, @, #, etc)
  
✅ **Exemplo correto:** `C:\Dados\contatos_origem\`  
❌ **Exemplo incorreto:** `C:\Dados\Contatos de Clientes!\`

#### 📁 **Pasta de Destino** (arquivos processados)
- Crie outra pasta para receber os **arquivos formatados**
- Esta pasta pode estar vazia inicialmente
- ⚠️ **IMPORTANTE:** Mesmas regras de nomenclatura acima

✅ **Exemplo correto:** `C:\Dados\contatos_processados\`  
❌ **Exemplo incorreto:** `C:\Dados\Planilhas (final)\`

---

### **2️⃣ CONFIGURAÇÃO DOS CAMINHOS**

Preencha os caminhos das pastas nas células abaixo:

| Célula Nomeada | Descrição | Exemplo |
|----------------|-----------|---------|
| **Local_Origem** | Caminho completo da pasta com arquivos originais | `C:\Dados\contatos_origem\` |
| **Local_Destino** | Caminho completo da pasta para arquivos processados | `C:\Dados\contatos_processados\` |

⚠️ **ATENÇÃO:**
- Sempre termine o caminho com `\` (barra invertida)
- Use o caminho completo (ex: `C:\Pasta\SubPasta\`)
- Não use caminhos de rede mapeados como letras (ex: `Z:\`)

---

### **3️⃣ PROCESSAMENTO DOS DADOS**

Após configurar os caminhos:

#### 🔵 **Botão: CONSOLIDAR DADOS**

Clique neste botão para iniciar o processamento. O sistema irá:

1. ✅ Abrir cada arquivo da **Pasta de Origem**
2. ✅ Criar tabelas formatadas em todas as abas
3. ✅ Adicionar coluna "origem" identificando o arquivo
4. ✅ Consolidar todos os dados em uma aba "Consolidado"
5. ✅ Salvar os arquivos processados na **Pasta de Destino**

⏱️ **Aguarde:** O processamento pode levar alguns minutos dependendo da quantidade de arquivos.

---

### **4️⃣ CORREÇÃO DE DADOS**

Se identificar necessidade de **corrigir alguma informação**:

1. 📂 Vá até a **Pasta de Destino**
2. ✏️ Abra o arquivo correspondente
3. 🔧 Faça as correções necessárias nas abas individuais
4. 💾 Salve o arquivo
5. 🔄 Retorne a esta planilha e clique em **"ATUALIZAR DADOS"**

#### 🔵 **Botão: ATUALIZAR DADOS**

Este botão reprocessa os dados já existentes na **Pasta de Destino**:
- Reconsolida todas as informações
- Atualiza a aba "Consolidado"
- Aplica as validações de telefone novamente

---

### **⚠️ AVISOS IMPORTANTES**

#### 🚨 **ATENÇÃO - ABA "CONSOLIDADO"**

> **Todas as alterações ou personalizações feitas diretamente na aba "Consolidado" serão PERDIDAS e SOBRESCRITAS ao atualizar os dados novamente.**

**O que fazer:**
- ✅ Faça correções nos **arquivos individuais** (Pasta de Destino)
- ✅ Depois atualize os dados
- ❌ **NÃO** edite diretamente a aba "Consolidado"

---

### **📊 ESTRUTURA DOS DADOS PROCESSADOS**

Cada arquivo processado terá:

- 📋 **Tabelas formatadas** em cada aba
- 🏷️ **Coluna "origem"** com o nome do arquivo fonte
- 📑 **Aba "Consolidado"** com todos os registros
- ✅ **Validação de telefones** aplicada

**Colunas padrão:**
- `mes` - Mês de referência
- `nome` - Nome do contato
- `telefone` - Número de telefone
- `origem` - Arquivo de origem

---

### **🔍 VALIDAÇÕES APLICADAS**

O sistema valida automaticamente os telefones:

| Status | Descrição |
|--------|-----------|
| ✅ **OK** | Telefone válido com DDD correto |
| ⚠️ **Falta DDD** | Número local sem DDD |
| ❌ **DDD inválido** | DDD não cadastrado na base |
| ❌ **DDI inválido** | Código de país incorreto |
| ❌ **Formato inválido** | Número fora dos padrões |

**Formatos gerados (apenas para status OK):**
- `telefone_normalizado`: 55 (99) 999999999
- `telefone_CRM`: +55 99 999999999
- `telefone_Bot`: 5599999999999

---

### **❓ SOLUÇÃO DE PROBLEMAS**

#### **Erro: "Pasta não encontrada"**
- Verifique se o caminho está correto
- Confirme que a pasta existe
- Certifique-se que terminou com `\`

#### **Erro: "Nenhum arquivo encontrado"**
- Confirme que há arquivos `.xls` ou `.xlsx` na pasta origem
- Verifique se os arquivos não estão corrompidos

#### **Processamento muito lento**
- Arquivos muito grandes podem demorar mais
- Feche outros programas para liberar memória
- Considere processar em lotes menores

#### **Dados não aparecem em "Consolidado"**
- Verifique se as abas têm as colunas: `mes`, `nome`, `telefone`
- Confirme que há dados preenchidos nos arquivos
- Certifique-se que as tabelas foram criadas corretamente

---

### **💡 DICAS E BOAS PRÁTICAS**

1. ✅ **Faça backup** dos arquivos originais antes de processar
2. ✅ **Mantenha os originais** - A pasta de origem não é alterada
3. ✅ **Teste com poucos arquivos** primeiro para validar
4. ✅ **Nomeie arquivos** de forma clara e organizada
5. ✅ **Revise os dados** após o processamento
6. ✅ **Use "Atualizar Dados"** após correções individuais

---

### **📞 SUPORTE**

Em caso de dúvidas ou problemas:
- Revise estas instruções cuidadosamente
- Verifique os exemplos de caminho
- Confirme que seguiu todos os passos
- Teste com um arquivo pequeno primeiro

---

## ✅ CHECKLIST RÁPIDO

Antes de clicar em "CONSOLIDAR DADOS", confirme:

- [ ] Pasta de origem criada e nomeada corretamente (sem espaços/caracteres especiais)
- [ ] Pasta de destino criada e nomeada corretamente (sem espaços/caracteres especiais)
- [ ] Arquivos Excel colocados na pasta de origem
- [ ] Célula "Local_Origem" preenchida com caminho completo (terminando com `\`)
- [ ] Célula "Local_Destino" preenchida com caminho completo (terminando com `\`)
- [ ] Backup dos arquivos originais realizado

**Tudo pronto? Clique em "CONSOLIDAR DADOS"!** 🚀