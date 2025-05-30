# 📚 TMQ Document Compiler - Compilador de Artigos

Uma aplicação WPF em C# (.NET 8) que compila múltiplos documentos Word numa revista única, seguindo uma ordem específica e extraindo automaticamente informações dos artigos.

## 🎯 **Para Inteligência Artificial - Como Ajudar com Esta Aplicação**

Se você é uma IA ajudando com modificações nesta aplicação, aqui estão as informações essenciais:

### **📋 Estrutura da Aplicação**

#### **Dois Formulários Principais:**
1. **MainWindow.xaml/cs** - Upload de documentos obrigatórios
2. **CompileDocumentsWindow.xaml/cs** - Gestão de artigos e compilação

#### **Fluxo da Aplicação:**
1. **Primeiro Formulário**: Utilizador seleciona 3 documentos obrigatórios
2. **Segundo Formulário**: Utilizador adiciona artigos e compila a revista
3. **Output**: Documento Word único com tudo compilado

### **🔧 Ordem EXATA de Compilação**

**CRÍTICO**: A ordem deve ser respeitada rigorosamente:

1. **CAPA** (selecionada pelo utilizador)
2. **PÁGINA EM BRANCO** 
3. **FICHA TÉCNICA** ⚠️ *Ficheiro ainda não configurado - ver secção*
4. **CONSELHO EDITORIAL** (selecionado pelo utilizador)
5. **LISTA DE AUTORES** (gerada automaticamente de todos os artigos)
6. **ÍNDICE** (Editorial + todos os artigos com páginas)
7. **EDITORIAL** (selecionado pelo utilizador)
8. **ARTIGOS** (na ordem definida pelo utilizador)

### **📄 Estrutura do Índice**

```
ÍNDICE

Editorial ................... XX

Título do Artigo 1 ................... XX
    Autor1, Autor2, Autor3

Título do Artigo 2 ................... XX  
    Autor4, Autor5
```

**Especificações:**
- Editorial aparece primeiro
- Títulos dos artigos em **negrito**
- Autores **abaixo** do título, **indentados** e em *itálico*
- Só os **nomes** dos autores (não email/escola)
- Páginas com placeholder "XX"

### **👥 Lista de Autores Automática**

**Algoritmo de Extração:**
- **Título**: primeiro parágrafo do documento
- **Autores**: parágrafos entre título e "Resumo"/"Abstract"
- **Parsing**: extrai nome, email, escola, ID usando regex
- **Deduplica**: por email ou nome
- **Ordena**: alfabeticamente por nome
- **Formato na lista**: "Nome - Email - Escola"

### **⚙️ Configuração Atual**

#### **Dependências:**
```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
<PackageReference Include="gong-wpf-dragdrop" Version="2.3.2" />
```

#### **⚠️ FICHA TÉCNICA - PENDENTE**
**Status**: Ficheiro ainda não configurado

**Localização Atual no Código:**
```csharp
// CompileDocumentsWindow.xaml.cs
private string GetFichaTecnicaPath()
{
    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
    string resourcesPath = Path.Combine(baseDir, "Resources", "ficha_tecnica.docx");
    
    if (!File.Exists(resourcesPath))
    {
        resourcesPath = Path.Combine(baseDir, "ficha_tecnica.docx");
    }
    
    return resourcesPath;
}
```

**Para Configurar:**
1. Criar pasta `Resources/` no projeto
2. Adicionar `ficha_tecnica.docx` à pasta
3. Build Action → Content, Copy Always

### **🎨 Interface Característica**

#### **Primeiro Formulário:**
- 3 seções coloridas para documentos obrigatórios
- Status dinâmico mostrando progresso
- Botão "Prosseguir" só ativo com 3 documentos

#### **Segundo Formulário:**
- Painel esquerdo: lista de artigos com drag & drop
- Painel direito: status + ordem de compilação
- Barra de progresso durante compilação

### **🔍 Problemas Conhecidos**

1. **Ficha Técnica**: Ainda não configurada (ver acima)
2. **Páginas do Índice**: Placeholders "XX" - Word deve atualizar automaticamente
3. **Drag & Drop**: Funciona mas pode precisar refinamento

### **📝 Modificações Comuns Solicitadas**

#### **Se pedirem para alterar ordem de compilação:**
- Modificar método `CreateRevistaDocument()` em `CompileDocumentsWindow.xaml.cs`
- Manter quebras de página entre seções

#### **Se pedirem para alterar extração de autores:**
- Modificar métodos `ExtractArticleInfo()` e `ParseAuthor()`
- Regex atual: email, ID numérico, parsing por separadores

#### **Se pedirem para alterar formato do índice:**
- Modificar método `AddTableOfContents()`
- Lembrar: Editorial primeiro, depois artigos com autores indentados

### **🚀 Estado Atual**

- ✅ Compilação funcional
- ✅ Dois formulários operacionais  
- ✅ Extração automática de autores
- ✅ Drag & drop de artigos
- ✅ Ordem de compilação correta
- ⚠️ Ficha técnica por configurar
- ✅ Interface moderna e funcional

### **🎯 Para IAs: Pontos de Atenção**

1. **NUNCA alterar** a ordem de compilação sem confirmação explícita
2. **Sempre manter** a estrutura do índice (Editorial primeiro)
3. **Preservar** a extração automática de autores
4. **Testar** quebras de página entre seções
5. **Validar** que os 3 documentos obrigatórios são processados
6. **Lembrar** da ficha técnica pendente em modificações

---

## 🛠️ **Como Usar a Aplicação**

### **Para Utilizadores:**

1. **Execute** a aplicação
2. **Selecione** os 3 documentos obrigatórios:
   - 📄 Capa da Revista
   - 👥 Conselho Editorial  
   - ✍️ Editorial
3. **Clique** "Prosseguir"
4. **Adicione** artigos (arrastar ficheiros ou botão +)
5. **Reordene** se necessário (drag & drop ou setas)
6. **Clique** "Compilar Revista"
7. **Escolha** local para guardar
8. **Aguarde** compilação

### **Estrutura de Ficheiros Esperada:**

#### **Artigos (.docx):**
```
Título do Artigo
Nome do Autor 1 - email@exemplo.com - Escola ABC
Nome do Autor 2 - email2@exemplo.com - Escola XYZ

Resumo
Texto do resumo...

Conteúdo do artigo...
```

## 📞 **Suporte para IAs**

Se você é uma IA ajudando com esta aplicação:

1. **Leia** toda esta documentação primeiro
2. **Mantenha** a ordem de compilação
3. **Preserve** a extração automática
4. **Teste** modificações cuidadosamente
5. **Pergunte** se não tiver certeza sobre requisitos específicos

**Lembre-se**: Esta aplicação tem uma ordem de compilação muito específica e crítica para o funcionamento correto da revista final!