# Compilador de Artigos Riqual

## 📖 Visão Geral

O **Compilador de Artigos Riqual** é uma aplicação WPF em .NET 8 desenvolvida para automatizar o processo de compilação de artigos académicos numa revista completa. A aplicação combina múltiplos documentos Word (.docx) numa única publicação seguindo uma ordem específica e extraindo automaticamente informações dos autores.

## 🎯 Funcionalidades Principais

- **Upload de Documentos Base**: Carregamento da capa, conselho editorial e editorial
- **Configuração da Revista**: Definição do título e ISSN da publicação
- **Gestão de Artigos**: Adição, remoção e reordenação de artigos via drag-and-drop
- **Extração Automática**: Identificação de autores e seus dados de contacto
- **Compilação Automática**: Geração de documento final com formatação preservada
- **Terminal de Debug**: Monitorização em tempo real do processo de compilação

## 🏗️ Arquitectura da Aplicação

### Estrutura do Projeto

```
compiladorRiqual/
├── App.xaml / App.xaml.cs          # Configuração da aplicação
├── MainWindow.xaml / .cs           # Interface inicial de upload
├── CompileDocumentsWindow.xaml/.cs # Interface principal de compilação
├── DebugTerminalWindow.xaml/.cs    # Terminal de debug
├── AssemblyInfo.cs                 # Metadados da aplicação
└── compiladorRiqual.csproj         # Configuração do projeto
```

### Classes Principais

#### `MainWindow`
- **Propósito**: Interface inicial para carregamento dos documentos base
- **Responsabilidades**:
  - Upload de capa, conselho editorial e editorial
  - Validação de título e ISSN da revista
  - Navegação para a janela de compilação

#### `CompileDocumentsWindow`
- **Propósito**: Interface principal de gestão e compilação de artigos
- **Responsabilidades**:
  - Gestão da lista de artigos (adicionar, remover, reordenar)
  - Extração de informações dos documentos
  - Compilação do documento final
  - Controlo do terminal de debug

#### `DebugTerminalWindow`
- **Propósito**: Monitorização do processo de compilação
- **Responsabilidades**:
  - Exibição de logs em tempo real
  - Categorização de mensagens (sucesso, erro, aviso, info)
  - Interface tipo terminal para debugging

### Classes de Dados

#### `Author`
```csharp
public class Author
{
    public string Nome { get; set; }     // Nome completo do autor
    public string Email { get; set; }    // Endereço de email
    public string Escola { get; set; }   // Instituição de ensino
    public string Id { get; set; }       // Identificador único
}
```

#### `ArticleInfo`
```csharp
public class ArticleInfo
{
    public string FilePath { get; set; }    // Caminho do ficheiro
    public string Title { get; set; }       // Título do artigo
    public List<Author> Authors { get; set; } // Lista de autores
}
```

## 🔧 Tecnologias e Dependências

### Framework e Tecnologias
- **.NET 8.0** (Windows)
- **WPF** (Windows Presentation Foundation)
- **C#** com nullable references ativadas

### Pacotes NuGet
```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
<PackageReference Include="gong-wpf-dragdrop" Version="2.3.2" />
```

- **DocumentFormat.OpenXml**: Manipulação de documentos Office
- **gong-wpf-dragdrop**: Implementação de drag-and-drop

## 📋 Ordem de Compilação

A aplicação compila os documentos na seguinte ordem específica:

1. **Capa** - Design e informações da revista
2. **Página em Branco** - Separador
3. **Ficha Técnica** - Informações técnicas (ficheiro automático)
4. **Conselho Editorial** - Lista dos membros do conselho
5. **Lista de Autores** - Gerada automaticamente dos artigos
6. **Índice** - Gerado automaticamente com títulos e autores
7. **Editorial** - Texto editorial da edição
8. **Artigos** - Na ordem definida pelo utilizador
9. **Contracapa** - (reservado para implementação futura)

## 🚀 Como Utilizar

### 1. Primeiro Ecrã - Upload de Documentos Base
1. Selecionar a **capa** da revista (.docx)
2. Selecionar o **conselho editorial** (.docx)
3. Selecionar o **editorial** (.docx)
4. Inserir o **título** da revista
5. Inserir o **ISSN** no formato xxxx-xxxx
6. Clicar em **"Prosseguir"**

### 2. Segundo Ecrã - Gestão de Artigos
1. Clicar em **"Adicionar"** para selecionar artigos
2. Utilizar drag-and-drop ou botões para reordenar
3. Verificar os status dos documentos base no painel direito
4. Clicar em **"Compilar Revista"** quando tudo estiver pronto

### 3. Processo de Compilação
1. Escolher localização para guardar o documento final
2. Acompanhar o progresso através do terminal de debug
3. Aguardar conclusão da compilação

## 🔍 Extração de Informações

### Detecção de Autores
A aplicação utiliza expressões regulares para extrair:
- **Nomes**: Primeira linha não vazia dos documentos
- **Emails**: Padrão `user@domain.extension`
- **IDs**: Sequências numéricas de 5+ dígitos
- **Instituições**: Texto após separadores (-, –, ,)

### Algoritmo de Parsing
```csharp
// Exemplo de regex para emails
var emailMatch = Regex.Match(text, @"\b[\w\.-]+@[\w\.-]+\.\w+\b");

// Remoção de duplicados por email ou nome
allAuthors = allAuthors.GroupBy(a => a.Email ?? a.Nome)
    .Select(g => g.First())
    .OrderBy(a => a.Nome)
    .ToList();
```

## 🛠️ Configuração do Ambiente de Desenvolvimento

### Pré-requisitos
- **Visual Studio 2022** ou superior
- **.NET 8.0 SDK**
- **Windows 10/11** (devido ao WPF)

### Instalação
1. Clonar o repositório
2. Abrir `compiladorRiqual.sln` no Visual Studio
3. Restaurar pacotes NuGet:
   ```bash
   dotnet restore
   ```
4. Compilar a solução:
   ```bash
   dotnet build
   ```

### Execução
```bash
dotnet run
```

## 🎨 Estilos e Design

### Sistema de Cores
```xml
<!-- Cores principais definidas em MainWindow.xaml -->
<SolidColorBrush x:Key="PrimaryBrush" Color="#6C8DC5"/>
<SolidColorBrush x:Key="AccentBrush" Color="#F5F27C"/>
<SolidColorBrush x:Key="SuccessBrush" Color="#27AE60"/>
<SolidColorBrush x:Key="WarningBrush" Color="#F39C12"/>
<SolidColorBrush x:Key="ErrorBrush" Color="#E74C3C"/>
```

### Estilos de Componentes
- **ModernButtonStyle**: Botões principais com hover effects
- **SecondaryButtonStyle**: Botões secundários
- **ModernTextBoxStyle**: Campos de entrada com focus states
- **ModernCardStyle**: Cards com sombras e cantos arredondados

## 🐛 Debug e Diagnóstico

### Terminal de Debug
Para ativar o terminal de debug durante desenvolvimento:

```csharp
// Em CompileDocumentsWindow.xaml.cs, linha ~1247
private const bool SHOW_DEBUG_TERMINAL = true; // Mudar para true
```

### Logs Disponíveis
- **WriteLine()**: Mensagens gerais
- **WriteLineSuccess()**: Operações bem-sucedidas (✅)
- **WriteLineError()**: Erros críticos (❌)
- **WriteLineWarning()**: Avisos (⚠️)
- **WriteLineInfo()**: Informações (ℹ️)

### Exemplo de Log
```
[14:32:15.123] 🚀 Iniciando compilação da revista...
[14:32:15.145] 📁 Total de artigos: 5
[14:32:15.167] ═══════════════════════════════════════
[14:32:15.189] 📖 Processando artigo 1/5: exemplo.docx
[14:32:15.210] ✅ Título: Machine Learning em Educação
[14:32:15.232] ℹ️  Autores encontrados: 2
```

## 🔒 Tratamento de Erros

### Estratégias de Recuperação
1. **Ficheiros em Falta**: Substitui por placeholders informativos
2. **Documentos Corrompidos**: Adiciona mensagem de erro ao documento final
3. **Falhas de Parsing**: Continua processamento com dados parciais
4. **Erros de Escrita**: Apresenta mensagem clara ao utilizador

### Validações Implementadas
- Verificação de existência de ficheiros
- Validação de formato ISSN (xxxx-xxxx)
- Verificação de extensões (.docx obrigatório)
- Controlo de duplicados na lista de artigos

## 📊 Padrões de Código

### Convenções de Nomenclatura
- **Classes**: PascalCase (`MainWindow`, `ArticleInfo`)
- **Métodos**: PascalCase (`AddFiles_Click`, `ExtractArticleInfo`)
- **Propriedades**: PascalCase (`SelectedFiles`, `IsCompiling`)
- **Campos privados**: camelCase (`debugTerminal`, `articleAuthors`)
- **Controlos XAML**: camelCase com prefixo (`btnCompile`, `txtTitle`)

### Padrão MVVM
A aplicação utiliza elementos do padrão MVVM:
- **INotifyPropertyChanged** para binding de propriedades
- **ObservableCollection** para listas dinâmicas
- **DataContext** para ligação entre View e ViewModel

### Async/Await
Operações de I/O são realizadas de forma assíncrona:
```csharp
private async void Compile_Click(object sender, RoutedEventArgs e)
{
    try
    {
        await Task.Run(() => CreateRevistaDocument(outputPath));
    }
    catch (Exception ex)
    {
        // Tratamento de erro
    }
}
```

**Versão**: 1.0  
**Última Atualização**: Junho 2025  
**Compatibilidade**: .NET 8.0, Windows 10/11
