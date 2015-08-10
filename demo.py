from o365 import O365ApiInterface

if __name__=="__main__":
  o = O365ApiInterface('your-email-here', 'your-password-here')
  
  # Getting unread Mails
  unread_mails = o.get_unread()