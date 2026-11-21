# Third-party notices

HMA – Harris Matrix Adapter is licensed under the MIT License (see `LICENSE`).
The following third-party work is used in HMA. Its licence requires that the copyright notice and the permission
notice are reproduced; this is done here and in `script/HMA_v1.0.2.py` (in front of the function `hme_layout()`).

## HME – Harris Matrix Editor

- Project: <https://github.com/oeai-dac/HME-HarrisMatrixEditor>
- Licence: MIT License, Copyright (c) 2026 oeai-dac
- Use in HMA: the function `hme_layout()` re-implements the *Auto-Layout* procedure of HME, so that the positions that
  HMA writes into the `.json` file for HME are the ones HME calculates itself. HMA also writes files in the `.json`
  format that HME reads. HMA contains no other code of HME. HME is a separate project; HMA is not affiliated with it.

```
MIT License

Copyright (c) 2026 oeai-dac

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Programs that HMA only supports

HMC – Harris Matrix Composer, Airtable, FileMaker, Microsoft Access, Google Sheets, PostgreSQL, MySQL and the other
programs named in the README are **not** part of HMA. HMA contains none of their code; the names belong to their
owners and are used only to describe compatibility.
