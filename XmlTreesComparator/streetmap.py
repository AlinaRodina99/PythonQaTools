import xmltodict

# fin = open('C:/Users/KS/Downloads/map1.osm', 'r', encoding='utf8')
# xml = fin.read()
# fin.close()
#
# dct_ = xmltodict.parse(xml)
# print(dct_)


fin = open('C:/Users/KS/Downloads/map2.osm', 'r', encoding='utf8')
xml = fin.read()
fin.close()
r = 0

dct_ext = xmltodict.parse(xml)
for node in dct_ext['osm']['node']:
    if 'tag' in node:
        tags = node['tag']
        if isinstance(tags, list):
            for tag in tags:
                if '@k' in tag and tag['@k'] == 'amenity' and tag['@v'] == 'fuel':
                    r +=1
        elif isinstance(tags, dict):
            if (tags['@v']) == 'fuel':
                r += 1
print(r)
for node in dct_ext['osm']['way']:
    if 'tag' in node:
        tags = node['tag']
        if isinstance(tags, list):
            for tag in tags:
                if '@k' in tag and tag['@k'] == 'amenity' and tag['@v'] == 'fuel':
                    r +=1
        elif isinstance(tags, dict):
            if (tags['@v']) == 'fuel':
                r += 1
print(r)
