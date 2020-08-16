import requests

def dataFetch():
  data = {
  'page': 1,
  'rows': 10,
  'shop_type': 0,
  'search_str': '美妆',
  'search_data': ["美妆"],
  }
  url = "https://z.vanmmall.com/index.php/index/goods"
  r = requests.post(url,params=data)
  print(r)
  result = r.json()
  print(result)

dataFetch()
