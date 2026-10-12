import re, sys
tpl, leaflet_css, basemap, legs, out_fragment, out_full = sys.argv[1:7]
s = open(tpl).read()
css = open(leaflet_css).read()
css = re.sub(r"url\(images/[^)]*\)", "none", css)
css = css.replace("behavior:url(#default#VML)", "").replace("behavior: url(#default#VML)", "")
s = s.replace("/*__LEAFLET_CSS__*/", css)
s = s.replace("/*__BASEMAP__*/null", open(basemap).read().strip())
s = s.replace("/*__LEGS__*/null", open(legs).read().strip())
assert "__" not in re.sub(r"__proto__", "", s.split("<script>")[0][-200:])
open(out_fragment, "w").write(s)
full = ('<!doctype html>\n<html lang="nl">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
        '<style>html,body{margin:0}[hidden]{display:none!important}</style>\n')
head_end = s.index("<div class=\"app\">")
full += s[:head_end] + "</head>\n<body>\n" + s[head_end:] + "\n</body>\n</html>\n"
open(out_full, "w").write(full)
print(len(s), len(full))
