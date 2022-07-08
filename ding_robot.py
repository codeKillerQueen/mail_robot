from dingtalkchatbot.chatbot import DingtalkChatbot

def ding(msg):
    webhook ='your dingding chat webhook'
    xiaoding = DingtalkChatbot(webhook)
    xiaoding.send_text(msg='[runner]' + '\n' + msg, at_mobiles=[your phone number])
