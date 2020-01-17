from apscheduler.schedulers.blocking import BlockingScheduler
# from datastatistics import query_brands_org_data, query_alogrithm_org_data
import logging
import time
import os
from apscheduler.triggers.cron import CronTrigger
from weekly import run

def initlogger(level):
    '''
    初始化日志配置
    '''
    # 第一步，创建一个logger
    logging.basicConfig()
    logger = logging.getLogger()
    logger.setLevel(level)  # Log等级总开关
    # 第二步，创建一个handler，用于写入日志文件
    rq = time.strftime('%Y%m%d%H%M', time.localtime(time.time())) + 'SchedulerLogs'
    log_path = os.path.dirname(os.getcwd()) + '/AcTimelyUpdateSchedulerLogs/'
    if not os.path.exists(log_path):
        os.mkdir(log_path)
    log_name = log_path + rq + '.log'
    logfile = log_name
    fh = logging.FileHandler(logfile, mode='w')
    fh.setLevel(level)  # 输出到file的log等级的开关
    # 第三步，定义handler的输出格式
    formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
    fh.setFormatter(formatter)
    # 第四步，将logger添加到handler里面
    logger.addHandler(fh)
    return logger


def scheduler():
    logger = initlogger(logging.INFO)
    try:
        logger.info('数据初始化完成')
    except Exception as e:
        logger.info('数据初始化异常: ' + e)
    scheduler = BlockingScheduler()
    day_insert_cron = CronTrigger(year='*', month='*', day='*', week='*', day_of_week='fir', hour='8', minute='0', second='0')
    scheduler.add_job(run,day_insert_cron)
    scheduler.start()

if __name__ == "__main__":
    scheduler()