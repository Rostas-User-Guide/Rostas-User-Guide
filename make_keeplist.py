#!/usr/bin/env python3
"""
Work out which images the guide still uses, so the rest can be removed.

    python3 make_keeplist.py          # write _keeplist.txt, print what is unused
    python3 make_keeplist.py --rm     # also `git rm` the unused images

2 Nov 2026 — the old hand-made _keeplist.txt went stale (it missed crops/ and
the originals crops.json crops from). This rebuilds it from the real sources
every time instead:
  - every image or video file named anywhere in index.html (src, data-full,
    poster, url(...)), which is also everything generate_docx.js embeds
  - every original screenshot crops.json crops from (make_crops.py needs them)
  - everything in crops/
Non-image files (scripts, HTML, the DOCX/PDF) are never touched.
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
MEDIA = re.compile(r'\.(png|jpe?g|gif|webp|svg|mp4)$', re.I)
NAMED = re.compile(r'''[\w./ -]+?\.(?:png|jpe?g|gif|webp|svg|mp4)(?=['")\s])''', re.I)


def tracked():
    out = subprocess.run(['git', 'ls-files', '-z'], cwd=HERE, capture_output=True, check=True)
    return [f for f in out.stdout.decode().split('\0') if f]


def main():
    page = open(os.path.join(HERE, 'index.html'), encoding='utf-8').read()
    keep = {m.strip().lstrip('./') for m in NAMED.findall(page)}
    crops = json.load(open(os.path.join(HERE, 'crops.json')))['crops']
    keep |= {c['src'] for c in crops.values()}
    keep |= {'crops/' + n + '.png' for n in crops}

    files = tracked()
    keep |= {f for f in files if f.startswith('crops/')}
    missing = sorted(f for f in keep if not os.path.exists(os.path.join(HERE, f)))
    # Untracked media is left alone: it may be new work not yet committed.
    unused = sorted(f for f in files
                    if f not in keep and (MEDIA.search(f) or '.' not in os.path.basename(f))
                    and os.path.exists(os.path.join(HERE, f)))

    with open(os.path.join(HERE, '_keeplist.txt'), 'w') as fh:
        fh.write('\n'.join(sorted(keep)) + '\n')

    size = sum(os.path.getsize(os.path.join(HERE, f)) for f in unused)
    print(f'{len(keep)} files kept (written to _keeplist.txt)')
    for f in missing:
        print(f'!! named but missing: {f}')
    print(f'{len(unused)} unused files, {size / 1e6:.0f} MB:')
    for f in unused:
        print('   ' + f)
    if '--rm' in sys.argv and unused:
        subprocess.run(['git', 'rm', '-q', '--'] + unused, cwd=HERE, check=True)
        print('Removed with git rm (still in git history if ever needed).')
    sys.exit(1 if missing else 0)


if __name__ == '__main__':
    main()
