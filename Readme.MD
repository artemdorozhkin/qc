# Библиотека QC (Quick Constructors) для VBA

Библиотека QC (Quick Constructors) предоставляет удобные функции для создания объектов в VBA. Эта библиотека упрощает создание различных объектов, таких как словари, файловые объекты, объекты работы с базой данных и многое другое. В этом README представлены примеры использования основных функций библиотеки.

### Установка

- Скачайте файлы библиотеки QC.
- В Excel, откройте Visual Basic for Applications (VBA) через меню "Разработчик" > "Visual Basic".

#### Стандартный способ

- Нажмите правой кнопкой мыши в окне VBAProject Explorer (View > VBAProject Explorer или Ctrl + R если оне не открыт) и выберите Import File... .
- В появившемся диалоговом окне найдите и выберите "QCScripting.bas" и "QCScriptingEnums" (или другие необходимые модули библиотеки).
- Нажмите "OK", чтобы добавить модули библиотеки в ваш проект.

#### Rubberduck v2

- Нажмите правой кнопкой мыши в окне Rubberduck Explorer и выберите Sync Project > Update Components from Files... .
- В появившемся диалоговом окне найдите и выберите "QCScripting.bas" и "QCScriptingEnums" (или другие необходимые модули библиотеки).
- Нажмите "OK", чтобы добавить модули библиотеки в ваш проект.

## Создание объектов библиотеки Scripting

### Создание объекта `Scripting.Dictionary`

```vb
Dim myDictionary As Object
Set myDictionary = NewDictionary(vbTextCompare)
```

### Создание объекта `Scripting.FileSystemObject`

```vb
Dim fso As Object
Set fso = NewFileSystemObject()
```

### Создание объекта `Scripting.Drive`

```vb
Dim drive As Object
Set drive = NewDrive("C:\")
```

### Создание объекта `Scripting.Folder`

```vb
Dim folder As Object
Set folder = NewFolder("C:\ExampleFolder")
```

### Создание объекта `Scripting.File`

```vb
Dim file As Object
Set file = NewFile("C:\ExampleFolder\example.txt")
```

### Создание объекта `Scripting.TextStream`

```vb
Dim textStream As Object
Set textStream = NewTextStream("C:\ExampleFolder\example.txt", qcIOModeForWriting, qcFormatForUnicode)
```

### Создание объекта `Scripting.StandardStream`

```vb
Dim standardStream As Object
Set standardStream = NewStandardStream(qcStandardStreamTypesForInput)
```

## Создание объектов библиотеки VBScript

### Создание объекта `RegExp`

```vb
Dim regex As Object
Set regex = NewRegExp(i:=True, g:=True, m:=False, Pattern:="\d{3}-\d{2}-\d{4}")
```

## Создание объектов библиотеки ADODB (ActiveX Data Objects)

### Создание объекта `ADODB.Connection`

```vb
Dim conn As Object
Set conn = NewConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;")
```

### Создание объекта `ADODB.Recordset`

```vb
Dim rs As Object
Set rs = NewRecordset()
```

### Создание объекта `ADODB.Command`

```vb
Dim cmd As Object
Set cmd = NewCommand()
```

### Создание объекта `ADODB.Parameter`

```vb
Dim param As Object
Set param = NewParameter()
```

### Создание объекта `ADODB.Record`

```vb
Dim record As Object
Set record = NewRecord()
```

### Создание объекта `ADODB.Stream`

```vb
Dim stream As Object
Set stream = NewStream()
```

## Создание объектов библиотеки WScript

### Создание объекта `WScript.Shell`

```vb
Dim shell As Object
Set shell = NewShell()
```

### Создание объекта `WScript.Network`

```vb
Dim network As Object
Set network = NewNetwork()
```

### Создание ярлыка с помощью `WScript.Shell`

```vb
Dim shortcut As Object
Set shortcut = NewShortCut("C:\Example.lnk", "C:\TargetFile.exe", qcStdIcons.WindowIcon)
```

## Создание объектов для приложений Microsoft Office

### Создание объекта `Outlook.Application`

```vb
Dim outlookApp As Object
Set outlookApp = NewOutlook()
```

### Создание объекта `Excel.Application`

```vb
Dim excelApp As Object
Set excelApp = NewExcel()
```

### Создание объекта `Word.Application`

```vb
Dim wordApp As Object
Set wordApp = NewWord()
```

### Создание объекта `Access.Application`

```vb
Dim accessApp As Object
Set accessApp = NewAccess()
```

### Создание объекта `PowerPoint.Application`

```vb
Dim powerPointApp As Object
Set powerPointApp = NewPowerPoint()
```

### Создание объекта `MSProject.Application`

```vb
Dim projectApp As Object
Set projectApp = NewMSProject()
```

### Создание объекта `Publisher.Application`

```vb
Dim publisherApp As Object
Set publisherApp = NewPublisher()
```

### Создание объекта `Visio.Application`

```vb
Dim visioApp As Object
Set visioApp = NewVisio()
```

### Создание объекта `Adobe Acrobat`

```vb
Dim acrobatApp As Object
Set acrobatApp = NewAdobeAcrobat()
```
