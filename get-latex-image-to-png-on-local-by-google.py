import urllib
query_url2 = urllib.parse.quote("\\Large x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}")
query_url = "http://chart.googleapis.com/chart?cht=tx&chl=" + query_url2
chart = urllib.request.urlopen(query_url)
f = open("mathf2.png","wb")
f.write(chart.read())
f.close()