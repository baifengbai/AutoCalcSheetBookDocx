import urllib.parse
import urllib.request
query_url2 = urllib.parse.quote(r"\LARGE A_s \geq \frac{V}{\alpha_r \alpha_v f_y} + \frac{N}{0.8 \alpha_b f_y} + \frac{M}{1.3 \alpha_r \alpha_b f_y z} ")
query_url = "http://latex.xuming.science/latex-image.php?math=" + query_url2
chart = urllib.request.urlopen(query_url)
f = open("mathf3.png","wb")
f.write(chart.read())
f.close()