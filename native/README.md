# Native TabloOkuma DLL source

This directory contains the C source supplied by the add-on author and its native build script. The separate source archive supplied x86 and x64 DLL files that are byte-for-byte identical to the binaries bundled with this add-on, as verified using SHA-256 checksums.

## Contents

- `tablo.c`, `tablo.h`: Windows native implementation and C ABI
- `tablo.def`: exports (`Tablo_ABI` and `Tablo_Satir`)
- `derle.bat`: MSVC build script
- `examples/`: optional examples in C, Python and C#
- `ORIGINAL_README_TR.txt`: author's original DLL documentation (Turkish)

## Build on Windows

Install Microsoft Visual Studio C++ Build Tools, including the Windows SDK. Open an **x64 Native Tools Command Prompt** for x64, or **x86 Native Tools Command Prompt** for x86. In this `native` folder run:

```bat
derle.bat
```

The script produces `TabloOkuma.dll` in the same directory; rebuild separately in each architecture's prompt. Do not overwrite the bundled binaries until testing the rebuilt output with NVDA and the Python integration.

The C ABI exposes `Tablo_ABI` (returns 1) and `Tablo_Satir`, which formats a row from UTF-16 inputs. See `tablo.h` and `ORIGINAL_README_TR.txt` for detailed parameters and limits.

## Corresponding packaged binaries (SHA-256)

- x64: `a989ca7bc7a7aba01c29b75279937dac4410ff85bba53289968d427695954a30`
- x86: `93be9d6cd25cbb7df7d0065cc48a449057478fbffe005ed2fbb774c11d63d186`

**Verification performed:** archived native DLL bytes matched the packaged addon DLL bytes exactly. `gcc -std=c11 -Wall -Wextra -fsyntax-only` successfully checked `tablo.c` on Linux. No Windows MSVC rebuild or real NVDA runtime test was performed here; a source archive match is not a reproducible-build guarantee.

**Licensing:** the author has confirmed the files are their own work and approved GNU GPL version 2 only (`GPL-2.0-only`) for the native source and distributed DLLs. See [`../LICENSE`](../LICENSE) and [`../COPYRIGHT.md`](../COPYRIGHT.md).
