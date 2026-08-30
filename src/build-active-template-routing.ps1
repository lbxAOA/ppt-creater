# Promote only the locally approved replacement templates that are already
# production-ready in the established validation registry. This registry is the
# routing authority after C:/PPT模板 cleanup.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater'
$source=Join-Path $repo 'production-template-library.json'
