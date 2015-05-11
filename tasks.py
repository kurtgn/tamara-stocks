from invoke import task, run
import pdb
@task
def build():
    print('dasdasd')

    result = run("nosetests -v", warn = True)
    pdb.set_trace()
