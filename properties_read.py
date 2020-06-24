class Properties:
  fileName = ''
  def __init__(self, fileName):
    self.fileName = fileName

  def getProperties(self,key):
    pro_file = open(self.fileName, 'r')
    for line in  pro_file:
      if line.find('=') > 0:
        lines = line.replace('\n', '').split('=')
        if lines[0].strip() == key:
          return lines[1].strip()



