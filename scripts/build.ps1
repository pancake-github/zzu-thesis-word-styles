param([string]$Version = 'v0.1.0-beta.1') # 文件名与 Release 标签使用同一版本。
$ErrorActionPreference = 'Stop' # 任一步失败都不继续打包。
$root = Split-Path -Parent $PSScriptRoot # 从脚本位置确定仓库目录。
$output = Join-Path $root 'artifacts' # 仅写入忽略的构建目录。
if (Test-Path -LiteralPath $output) { throw 'artifacts 已存在，请先保留或移走旧构建结果，再开始新构建。' } # 不混入旧文件。
New-Item -ItemType Directory -Path $output | Out-Null # 创建本次构建目录。
Push-Location $root # 相对项目路径以仓库为基准。
try {
    dotnet restore tests/ZZUWordStyles.Core.Tests --locked-mode # 固定核心测试依赖。
    if ($LASTEXITCODE -ne 0) { throw '依赖校验失败。' } # 锁文件改变需要先经过审查。
    dotnet run --project tests/ZZUWordStyles.Core.Tests -c Release --no-restore # 运行实际核心测试。
    if ($LASTEXITCODE -ne 0) { throw '核心测试失败。' } # 测试失败禁止发布。
    $gui = Join-Path $output 'ZZUWordStyles' # ZIP 内包含一个完整应用目录。
    dotnet publish src/ZZUWordStyles.App/ZZUWordStyles.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -p:PublishTrimmed=false -p:PublishAot=false -p:DebugType=None -p:DebugSymbols=false -o $gui # 用户无需额外运行时。
    if ($LASTEXITCODE -ne 0) { throw '图形助手构建失败。' } # 不将部分输出误认为安装包。
    foreach ($name in @('README.md','LICENSE','THIRD-PARTY-NOTICES.md','CHANGELOG.md')) { Copy-Item -LiteralPath (Join-Path $root $name) -Destination $gui } # 随包交付说明与许可证。
    Copy-Item -LiteralPath (Join-Path $root 'docs') -Destination $gui -Recurse # 完整离线说明。
    Copy-Item -LiteralPath (Join-Path $root 'assets') -Destination $gui -Recurse # 可单独使用的已脱敏模板。
    $commit = (git rev-parse HEAD).Trim() # 记录这批构建对应的源码。
    @{Version=$Version;SourceCommit=$commit;Sdk=(dotnet --version);Runtime='win-x64';SelfContained=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $gui 'BUILD.json') -Encoding utf8 # 不记录本机用户名或路径。
    $template = Join-Path $output 'templates' # 模板专用包无需运行程序。
    New-Item -ItemType Directory -Path $template | Out-Null # 不读取个人 Normal。
    Copy-Item -LiteralPath (Join-Path $root 'assets') -Destination $template -Recurse # 仅包含公共模板与规格。
    Copy-Item -LiteralPath (Join-Path $root 'docs') -Destination $template -Recurse # 带上样式表和使用说明。
    foreach ($name in @('README.md','LICENSE','THIRD-PARTY-NOTICES.md')) { Copy-Item -LiteralPath (Join-Path $root $name) -Destination $template } # 随包说明使用边界。
    Compress-Archive -LiteralPath $gui -DestinationPath (Join-Path $output "ZZUWordStyles-$Version-win-x64.zip") -CompressionLevel Optimal # 打包完整自包含目录。
    Compress-Archive -LiteralPath $template -DestinationPath (Join-Path $output "ZZUWordStyles-$Version-templates.zip") -CompressionLevel Optimal # 单独提供小体积模板包。
    $lines = Get-ChildItem -LiteralPath $output -Filter '*.zip' | Sort-Object Name | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(),$_.Name } # 对最终 ZIP 计算摘要。
    $lines | Set-Content -LiteralPath (Join-Path $output 'SHA256SUMS.txt') -Encoding ascii # 所有下载件使用相同校验清单。
} finally { Pop-Location } # 恢复调用者工作目录。
