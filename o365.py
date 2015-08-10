# Author: ashrith.sheshan@gmail.com

import requests
import json
O365_URL = "https://outlook.office365.com"

class O365Exception(Exception):
  pass

class O365ApiInterface(object):

  def __init__(self, username, password):
    self.username = username
    self.password = password
    self.session = requests.session()

  def get_unread(self):
    url = "%s/api/v1.0/me/messages?$filter=IsRead eq false" % (
        O365_URL
    )
    result = self.session.get(url, auth=(self.username, self.password))
    if result.status_code!=200:
      raise O365Exception("Error: Cannot fetch mails")
    mails = result.json()['value']
    return {'success': True, 'total': len(mails), 'data': mails}

  def mark_as_read(self, message_id):
    url = "%s/api/v1.0/me/messages/%s" % (
        O365_URL, message_id
    )
    headers = {'Content-Type': 'application/json'}
    data = {'IsRead': True}
    result = self.session.patch(url, data= json.dumps(data),
                                headers= headers,
                           auth=(self.username, self.password))
    if result.status_code != 200:
      success = False
    else:
      success = True
    return {'success': success}



