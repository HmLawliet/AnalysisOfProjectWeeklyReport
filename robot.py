'''
    钉钉 发送文件至工作通知中
'''
from functools import wraps
import logging
import json
import requests


__slots__ = ['send_work_notice_file']


class Config:
    corpid = '*********'  # Nails are issued to the unique ID of each company
    agent_id = 11111111  # ID of the applet created in the nail background
    appkey = '*********'  # The key of the applet
    appsecret = 'c5VNrT4dKQUIv0ZjVQjQ9Ma8Q86gEuOkJnZonbQaldGvWk7Cx4dnK9vFFtNMB_-I'  # The secret of the applet
    userid_list = '*********'  # more user  format '*********,*********'
    sender = '*********'  # Sender of message
    cid = ''  #  Session id   It can be an individual or a group.  会话id（该cid和服务端开发文档-普通会话消息接口配合使用，而且只能使用一次，之后将失效）
    chatid = '*********'  # Group session id  会话id（该会话cid永久有效）


def choice_file_type(file, media_id=None, content=None, duration=None, messageUrl=None, picUrl=None,
                     title=None, text=None, bgcolor="FFBBBBBB", form=[], rich={}, file_count=None,
                     author=None, markdown={}, btn_orientation='1', btn_json_list=[], single_title=None,
                     single_url=None):
    '''
    For more information, please refer to the official nail development document
    url: https://open-doc.dingtalk.com/microapp/serverapi2/gh60vz.

    param file: the type of file    type str
    param media_id: Media_id is obtained by uploading the file to   type str
        "https://oapi.dingtalk.com/media/upload?Access_token=ACCESS_TOKEN&type=TYPE", 
        where ACCESS_TOKEN is obtained by "https://oapi.dingtalk.com/gettoken?Appkey=key&appsecret=secret".
        And appkey and appsecret need to create a small program to get through nail development background.
    param content: Content displayed    type str
    param duration: Audio Length    type str
    param messageUrl: Links to information  type str
    param picUrl: Links to picture   type str
    param title:  Main title  type str
    param text:  Subtitle  type str
    param bgcolor:  background color  type str
    param form: Show information   type list  example [{"key": "姓名:","value": "hm"},{"key": "姓名:","value": "hm"},],
    param rich: Show information   type  dict example {"num": "15.6","unit": "元"},
    param file_count: The count of file   type str  example  '1'
    param author: The information of author  type str  example 'hm'
    param markdown: Show information    type dict  example {"title": '中国最棒', "text": '最可爱的人是中国人',}
    param btn_orientation: When using the independent jump ActionCard style, the buttons are arranged 
        vertically (0) and horizontally (1); they must be set at the same time as btn_json_list.  example "1",
    param btn_json_list:A list of buttons when using the independent jump ActionCard style; 
        must be set at the same time as btn_orientation   
        example [{"title": "一个按钮","action_url": "https://www.taobao.com"},{"title": "两个按钮","action_url": "https://www.tmall.com"}]
    param single_title       
    '''
    if file == 'file':
        return {
            "msgtype": "file",
            "file": {
                "media_id": media_id,
            }
        }
    elif file == 'text':
        return {
            "msgtype": "text",
            "text": {
                "content": content,
            }
        }
    elif file == 'image':
        return {
            "msgtype": "image",
            "image": {
                "media_id": media_id,
            }
        }
    elif file == 'voice':
        return {
            "msgtype": "voice",
            "voice": {
                "media_id": media_id,
                "duration": duration,
            }
        }
    elif file == 'link':
        return {
            "msgtype": "link",
            "link": {
                "messageUrl": messageUrl,
                "picUrl": picUrl,
                "title": title,
                "text": text
            }
        }
    elif file == 'oa':
        return {
            "msgtype": "oa",
            "oa": {
                "message_url": messageUrl,
                "head": {
                    "bgcolor": bgcolor,
                    "text": text
                },
                "body": {
                    "title": title,
                    "form": form,
                    "rich": rich,
                    "content": content,
                    "image": media_id,
                    "file_count": file_count,
                    "author": author,
                }
            }
        }
    elif file == 'markdown':
        return {
            "msgtype": "markdown",
            "markdown": markdown,
        }
    elif file == 'action_card':
        if not single_url and not single_title:
            # 独立跳转
            return {
                "msgtype": "action_card",
                "action_card": {
                    "title": title,
                    "markdown": text or markdown,
                    "btn_orientation": btn_orientation,
                    "btn_json_list": btn_json_list,
                }
            }
        else:
            # 整体跳转
            return {
                "msgtype": "action_card",
                "action_card": {
                    "title": title,
                    "markdown": text or markdown,
                    "single_title": single_title,
                    "single_url": single_url,
                }
            }


def tryExcept(func):
    '''Capture exception information'''
    @wraps(func)
    def wrapper(file):
        try:
            return func(file)
        except Exception as e:
            logging.warning(e)
            return False
    return wrapper


def get_value(response,key):
    '''Getting values from data in JSON format'''
    try:
        res = json.loads(response.text)
    except Exception as e:
        raise e
    return res[key]


@tryExcept
def get_media_file_id(file):
    # get access_token 
    access_token = requests.get(
        'https://oapi.dingtalk.com/gettoken?appkey=%s&appsecret=%s' % (Config.appkey, Config.appsecret))
    access_token = get_value(access_token,'access_token')
    # get media_id 
    upload_url = 'https://oapi.dingtalk.com/media/upload?access_token=%s&type=file' % (access_token)
    media_id = requests.post(upload_url, files={'media': open(file, 'rb')})
    # print(media_id.text)
    media_id = get_value(media_id,'media_id')
    return access_token,media_id


@tryExcept
def send_work_notice_file(file,file_type='image'):
    ''' 
    send work notice
    return bool :  True is successed / False is failed or error
    '''
    # get access_token  media_id
    res = get_media_file_id(file)
    if not res:
        raise Exception('get access_token or get media_id has error')
    access_token,media_id = res
    # send work notice
    send_url = 'https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2?access_token=%s' % (
        access_token)
    data = {
        'agent_id': Config.agent_id,
        'userid_list':Config.userid_list,
        'msg': choice_file_type(file_type,media_id)
    }
    res = requests.post(send_url, data=json.dumps(data))
    return get_value(res,'errcode') == 0


@tryExcept
def send_common_information_file(file):
    '''
    send common information
    return bool : True is successed / False is failed or error
    '''
    file_type='file'
    # get access_token  media_id
    res = get_media_file_id(file)
    if not res:
        raise Exception('get access_token or get media_id has error')
    access_token,media_id = res
    # send common information
    send_url = 'https://oapi.dingtalk.com/message/send_to_conversation?access_token=%s' % access_token
    data = {
        'sender':Config.sender,
        'cid':Config.cid,
        'msg':choice_file_type(file_type,media_id)
    }
    res = requests.post(send_url, data=json.dumps(data))
    return get_value(res,'errcode') == 0

@tryExcept
def send_group_messages_file(file):
    '''
    send group messages
    return bool : True is successed / False is failed or error
    '''
    file_type='file'
    # get access_token  media_id
    res = get_media_file_id(file)
    if not res:
        raise Exception('get access_token or get media_id has error')
    access_token,media_id = res 
    # send group messages 
    send_url = 'https://oapi.dingtalk.com/chat/send?access_token=%s' % access_token
    data = {
        'access_token':access_token,
        'chatid':Config.chatid, # 群会话id
        'msg':choice_file_type(file_type,media_id)
    }
    res = requests.post(send_url, data=json.dumps(data))
    # print(res.text)
    return get_value(res,'errcode') == 0



if __name__ == "__main__":
    file = 'pws.png'
    send_work_notice_file(file)
