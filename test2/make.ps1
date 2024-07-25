$vb98_path = Join-Path ${Env:ProgramFiles(x86)} "Microsoft Visual Studio\VB98"
$vb6_exe = Join-Path $vb98_path "VB6.EXE"

$build_dir = ".\build"

$vbp_file = "RegexTest2Project.vbp"

if (!(Test-Path $build_dir)) { [void]( New-Item $build_dir -ItemType Directory ) }

# build project
#   Linking to console application specified in vbp file, as described here:
#      https://bbs-vbstreets-ru.translate.goog/viewtopic.php?f=28&t=44358&_x_tr_sch=http&_x_tr_sl=ru&_x_tr_tl=en&_x_tr_hl=ru
#   piping to Out-Null required for script to continue only after completion
& $vb6_exe /MAKE $vbp_file /outdir $build_dir | Out-Null
