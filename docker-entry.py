from apscheduler.schedulers.blocking import BlockingScheduler
from subprocess import call

def job():
    call(["sh", "service-now-generate-tracker.sh"])

scheduler = BlockingScheduler()
scheduler.add_job(job, 'interval', minutes=1)
scheduler.start()
