[CmdletBinding()]
param([string]$RootPath)
$ErrorActionPreference = 'Stop'
$nativeSource = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

internal static class HovrinoFolders
{
    private static readonly Dictionary<string,string> Albums = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase) {
        { "R2.1", "01. Генеральный план\\63003-1-26-ХВР-РД-ГП - Генеральный план" },
        { "R2.2", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР1 - Планы, разрезы, фасады" },
        { "R2.3", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР2 - Окна, двери, ворота" },
        { "R2.4", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР3 - Схемы расположения сэндвич-панелей фасадов" },
        { "R2.5", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР4 - Схемы расположения сэндвич-панелей внутренних стен и перегородок" },
        { "R2.6", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР5 - Детали полов, кровли, стен и перегородок" },
        { "R2.7", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР6 - Узлы стен из сэндвич-панелей" },
        { "R2.8", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР7 - Планы с привязками перегородок" },
        { "R2.9", "02. Архитектурные решения\\63003-1-26-ХВР-РД-1-АР8 - Внутренняя отделка. Спецификации. Ведомости" },
        { "R2.10", "02. Архитектурные решения\\63003-1-26-ХВР-РД-2-АР - Архитектурные решения. Вспомогательные здания и сооружения" },
        { "R2.13", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ1 - Фундаменты. Здания склада" },
        { "R2.14", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ2 - Фундаменты. Вспомогательные здания и сооружения" },
        { "R2.15", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ3 - Фундаменты оборудования (доки)" },
        { "R2.16", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ4 - Цокольные панели" },
        { "R2.17", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ4И - Цокольные панели. Изделия" },
        { "R2.18", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ5 - Колонны" },
        { "R2.19", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ5И - Колонны. Изделия" },
        { "R2.20", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ6 - Перекрытия" },
        { "R2.21", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ7 - Лестницы внутренние" },
        { "R2.22", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ7И - Лестницы внутренние. Изделия" },
        { "R2.23", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ8 - Плиты полов и приямки" },
        { "R2.24", "03. Конструкции железобетонные\\63003-1-26-ХВР-РД-1-КЖ9 - Отбойники (направляющие колёс)" },
        { "R2.25", "04. Конструкции металлические\\63003-1-26-ХВР-РД-1-КМ1 - Покрытие" },
        { "R2.27", "04. Конструкции металлические\\63003-1-26-ХВР-РД-1-КМ3 - Фахверк наружный - внутренний" },
        { "R2.28", "04. Конструкции металлические\\63003-1-26-ХВР-РД-1-КМ4 - Металлические лестницы (наружные и внутренние)" },
        { "R2.29", "04. Конструкции металлические\\63003-1-26-ХВР-РД-1-КМ5 - Ограждения и защитные конструкции" },
        { "R2.30", "04. Конструкции металлические\\63003-1-26-ХВР-РД-2-КМ - Конструкции металлические крытой парковки" },
        { "R2.31", "05. Электроснабжение\\63003-1-26-ХВР-РД-1-ЭОМ1 - Внутреннее электроосвещение и силовое электрооборудование. Здания склада" },
        { "R2.32", "05. Электроснабжение\\63003-1-26-ХВР-РД-1-ЭОМ2 - Внутреннее электроосвещение и силовое электрооборудование. Вспомогательные здания и сооружения" },
        { "R2.33", "05. Электроснабжение\\63003-1-26-ХВР-РД-1-ЭМ - Силовое электрооборудование, молниезащита и заземление" },
        { "R2.34", "06. Водоснабжение и канализация\\63003-1-26-ХВР-РД-1-ВК - Внутренние сети водоснабжения и водоотведения. Здания склада" },
        { "R2.35", "06. Водоснабжение и канализация\\63003-1-26-ХВР-РД-2-ВК - Внутренние сети водоснабжения и водоотведения. Вспомогательные здания и сооружения" },
        { "R2.37", "07. Отопление, вентиляция и кондиционирование\\63003-1-26-ХВР-РД-1-ОВ1 - Отопление" },
        { "R2.38", "07. Отопление, вентиляция и кондиционирование\\63003-1-26-ХВР-РД-1-ОВ2 - Вентиляция общеобменная. Вентиляция противодымная" },
        { "R2.39", "07. Отопление, вентиляция и кондиционирование\\63003-1-26-ХВР-РД-1-ОВ3 - Кондиционирование" },
        { "R2.41", "08. Теплоснабжение (тепломеханические системы)\\63003-1-26-ХВР-РД-1-ТМ - Индивидуальный тепловой пункт. Зданич склада" },
        { "R2.42", "09. Слаботочные противопожарные системы\\63003-1-26-ХВР-РД-1-АПС - Система пожарной сигнализации" },
        { "R2.43", "09. Слаботочные противопожарные системы\\63003-1-26-ХВР-РД-1-СОУЭ - Система оповещения и управления эвакуацией людей при пожаре" },
        { "R2.44", "09. Слаботочные противопожарные системы\\63003-1-26-ХВР-РД-1-ППА - Система противопожарной автоматики" },
        { "R2.45", "09. Слаботочные противопожарные системы\\63003-1-26-ХВР-РД-2-ППС - Система пожарной сигнализации. Система оповещения и управления эвакуацией людей при пожаре. Система противопожарной автоматики. Вспомогательные зданияи сооружения" },
        { "R2.46", "10. Системы внутреннего автоматического пожаротушения\\63003-1-26-ХВР-РД-1-АУВПТ - Автоматическая установка водяного пожаротушения. Внутренний противопожарный водопровод" },
        { "R2.47", "10. Системы внутреннего автоматического пожаротушения\\63003-1-26-ХВР-РД-1-АУГПТ - Автоматическая установка газового пожаротушения" },
        { "R2.48", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-1-АОВ - Автоматизация вентиляции" },
        { "R2.49", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-1-АТМ - Автоматизация Индивидуального теплового пункта" },
        { "R2.50", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-1-АВО - Автоматизация воздушного отопления" },
        { "R2.51", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-1-АДИС - Автоматизация и диспетчерзация инженерных систем" },
        { "R2.52", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-1-АСТУЭ - Автоматизированная система технического учета энергоресурсов" },
        { "R2.53", "11. Автоматизация и диспетчерезация инженерых систем\\63003-1-26-ХВР-РД-2-АДИС - Автоматика, диспетчаризация инженерных систем" },
        { "R2.55", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-НЭС - Наружные внутриплощадочные сети электроснабжения 6 - 10 - 20-0,4 кВ с устройством ТП" },
        { "R2.56", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-НЭО - Наружное электроосвещение" },
        { "R2.57", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-НВ - Наружные сети водоснабжения" },
        { "R2.58", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-НК - Наружные сети канализации" },
        { "R2.59", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-ТС - Наружные тепловые сети" },
        { "R2.60", "12. Наружные внутриплощадочные сети\\63003-1-26-ХВР-РД-НСС - Наружные сети связи" },
        { "R2.61", "13. Прочее\\63003-1-26-ХВР-РД-ОМК - Огнезащита металлических конструкций" },
    };
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern uint GetFileAttributesW(string path);
    [DllImport("shell32.dll", CharSet=CharSet.Unicode)]
    private static extern int SHParseDisplayName(string name, IntPtr context, out IntPtr item, uint attributes, out uint resultAttributes);
    [DllImport("shell32.dll")]
    private static extern int SHOpenFolderAndSelectItems(IntPtr folder, uint count, IntPtr children, uint flags);

    [DllImport("ole32.dll")]
    private static extern int CoInitializeEx(IntPtr reserved,uint mode);
    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    [STAThread]
    private static void Main(string[] args)
    {
        bool comInitialized=false;
        try {
            if (args.Length != 1) throw new ArgumentException("Откройте комплект значком папки в графике Ховрино.");
            Match match=Regex.Match(args[0], @"\Anso-hovrino://album/(R2\.[0-9]+)/?\z", RegexOptions.IgnoreCase|RegexOptions.CultureInvariant);
            string relative;
            if (!match.Success || !Albums.TryGetValue(match.Groups[1].Value,out relative))
                throw new ArgumentException("Ссылка не соответствует комплекту РД Ховрино.");
            string root=File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"root.txt")).Trim();
            if (!Regex.IsMatch(root,@"\A(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))") || root.StartsWith(@"\\?\") || root.StartsWith(@"\\.\"))
                throw new ArgumentException("Повторно выполните настройку и укажите полный путь к папкам РД.");
            root=Path.GetFullPath(root).TrimEnd('\\');
            string folder=Path.GetFullPath(Path.Combine(root,relative));
            if (!folder.StartsWith(root+"\\",StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Папка находится вне каталога проекта.");
            string extended=folder.StartsWith(@"\\")?@"\\?\UNC\"+folder.Substring(2):@"\\?\"+folder;
            uint attributes=GetFileAttributesW(extended);
            if (attributes==UInt32.MaxValue || (attributes&16)==0)
                throw new IOException("Папка недоступна. Проверьте подключение сетевого диска или VPN, права доступа и создание папок по скрипту.\n\n"+folder);
            Marshal.ThrowExceptionForHR(CoInitializeEx(IntPtr.Zero,2));comInitialized=true;
            IntPtr item=IntPtr.Zero;uint resultAttributes;
            int result=SHParseDisplayName(folder,IntPtr.Zero,out item,0,out resultAttributes);
            if (result<0) {
                if (item!=IntPtr.Zero) { Marshal.FreeCoTaskMem(item);item=IntPtr.Zero; }
                result=SHParseDisplayName(extended,IntPtr.Zero,out item,0,out resultAttributes);
            }
            try {
                Marshal.ThrowExceptionForHR(result);
                Marshal.ThrowExceptionForHR(SHOpenFolderAndSelectItems(item,0,IntPtr.Zero,0));
            } finally { if(item!=IntPtr.Zero) Marshal.FreeCoTaskMem(item); }
        } catch(Exception ex) {
            MessageBox.Show(ex.Message,"Папки РД Ховрино",MessageBoxButtons.OK,MessageBoxIcon.Information);
        } finally { if(comInitialized)CoUninitialize(); }
    }
}

'@
try {
    $defaultRoot = 'P:\2026\02_Ховрино\07 Стадия Р\02. Разделы'
    if ([string]::IsNullOrWhiteSpace($RootPath)) {
        Write-Host 'Корневая папка РД:'
        Write-Host $defaultRoot
        $RootPath = Read-Host 'Enter - оставить путь, либо введите свой путь'
        if ([string]::IsNullOrWhiteSpace($RootPath)) { $RootPath = $defaultRoot }
    }
    $RootPath = $RootPath.Trim().Trim('"')
    if ($RootPath -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))' -or $RootPath.StartsWith('\\?\') -or $RootPath.StartsWith('\\.\')) {
        throw 'Укажите полный путь к диску или сетевой папке, например P:\Проект\РД или \\server\share\РД.'
    }
    $RootPath = [IO.Path]::GetFullPath($RootPath).TrimEnd('\')
    $installDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'NSO\HovrinoFolders'
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    $exe = Join-Path $installDir 'HovrinoFolders.exe'
    $tempExe = Join-Path $installDir ('HovrinoFolders-' + [Guid]::NewGuid().ToString('N') + '.exe')
    try {
        Add-Type -TypeDefinition $nativeSource -ReferencedAssemblies 'System.Windows.Forms.dll' -OutputAssembly $tempExe -OutputType WindowsApplication
        Move-Item -LiteralPath $tempExe -Destination $exe -Force
    } finally { if(Test-Path -LiteralPath $tempExe){Remove-Item -LiteralPath $tempExe -Force} }
    [IO.File]::WriteAllText((Join-Path $installDir 'root.txt'),$RootPath,(New-Object Text.UTF8Encoding($false)))
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\Classes\nso-hovrino')
    try {
        $key.SetValue('','URL:Папки РД Ховрино')
        $key.SetValue('URL Protocol','')
        $icon = $key.CreateSubKey('DefaultIcon')
        try { $icon.SetValue('',(Join-Path $env:WINDIR 'explorer.exe') + ',0') } finally { $icon.Dispose() }
        $command = $key.CreateSubKey('shell\open\command')
        try { $command.SetValue('',('"' + $exe + '" "%1"')) } finally { $command.Dispose() }
    } finally { $key.Dispose() }
    Write-Host 'Настройка выполнена. Вернитесь к графику и нажмите значок папки.'
    Write-Host ('Каталог: ' + $RootPath)
} catch {
    Write-Error $_.Exception.Message -ErrorAction Continue
    exit 1
}
