from dingtalkchatbot.chatbot import DingtalkChatbot

def ding(msg):
    webhook ='https://oapi.dingtalk.com/robot/send?access_token=e4d08e96c1dc58e09b22f18e5805635c617588829858c628412be35204889db4'
    xiaoding = DingtalkChatbot(webhook)
    xiaoding.send_text(msg='[runner]' + '\n' + msg, at_mobiles=[18280098532])
