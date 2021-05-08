import os

if os.name=='posix':
  from pyvirtualdisplay import Display

from functools import wraps

def virtual_display(func):
  @wraps(func)
  def wrapper():
    display = None
    if os.name=='posix':
      display = Display(visible=0, size=(800, 600))
      display.start()

    func()

    if display:
      display.stop()

  return wrapper
