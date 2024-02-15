# html_str = ('<html><body><table>'
#             '<tr><td>1</td><td>2</td></tr>'
#             '<tr><td>3</td><td>4</td></tr>'
#             '</table></body></html>')
# http://<это число>.ru  <a href=http://hse.ru>Высшая школа экономики</a>
html_str = '<html><body><table>'
for i in range(1, 11):
    html_str += '<tr>'
    for k in range(1, 11):
        html_str += ('<td>' + '<a href=http://' + str(i * k) + '.ru>'+str(i * k) + '</a>'+' '+'</td>')
    html_str += '</tr>'
html_str += '</table></body></html>'
Html_file = open("C:/Users/KS/Downloads/filename.html", "w")
Html_file.write(html_str)
Html_file.close()
