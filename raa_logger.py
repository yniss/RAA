import logging
import queue
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter import N, S, E, W

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)
logger_name = logger.name #TODO: consider passing to other processes, so he can call  "logger = logging.getLogger(logger_name)"

##  #TODO: figure out a way - where to add a logger file handler for all messages (including debug level)
##  # Create file handler
##  f_handler = logging.FileHandler(file_name='raa.log', mode='w')
##  f_handler.setLevel(logging.DEBUG)
##  
##  # Create formatter and add it to handler
##  f_format = logging.Formatter('%(asctime)s %(name)s [%(levelname)s]: %(message)s')
##  f_handler.setFormatter(f_format)
##  
##  # Add handler to the logger
##  logger.addHandler(f_handler)


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    It can be used from different threads
    The ConsoleUi class polls this queue to display records in a ScrolledText widget
    """
    # Example from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06
    # (https://stackoverflow.com/questions/13318742/python-logging-to-tkinter-text-widget) is not thread safe!
    # See https://stackoverflow.com/questions/43909849/tkinter-python-crashes-on-new-thread-trying-to-log-on-main-thread

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


class ConsoleUi:
    """Poll messages from a logging queue and display them in a scrolled text widget"""

    def __init__(self, frame, event):
        self.frame = frame
        self.event = event
        # Create a ScrolledText wdiget
        self.scrolled_text = ScrolledText(frame, state='disabled', height=24)
        self.scrolled_text.grid(row=0, column=0, sticky=(N, S, W, E)) 
        self.scrolled_text.configure(font='TkFixedFont')
        self.scrolled_text.tag_config('INFO', foreground='black')
        self.scrolled_text.tag_config('DEBUG', foreground='gray')
        self.scrolled_text.tag_config('WARNING', foreground='orange')
        self.scrolled_text.tag_config('ERROR', foreground='red')
        self.scrolled_text.tag_config('CRITICAL', foreground='red', underline=1)
        # Create a logging handler using a queue
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        formatter = logging.Formatter('%(message)s')
        self.queue_handler.setFormatter(formatter)
        logger.addHandler(self.queue_handler)
        # Start polling messages from the queue
        self.frame.after(100, self.poll_log_queue)

    def display(self, record):
        msg = self.queue_handler.format(record)
        self.scrolled_text.configure(state='normal')
        self.scrolled_text.insert(tk.END, msg + '\n', record.levelname)
        self.scrolled_text.configure(state='disabled')
        # Autoscroll to the bottom
        self.scrolled_text.yview(tk.END)

    def poll_log_queue(self):
        # Check every 100ms if there is a new message in the queue to display
        while True:
            try:
                record = self.log_queue.get(block=False)
            except queue.Empty:
                break
            else:
                self.display(record)

            event_set = self.event.is_set() #TODO: event polling does not work...
            if event_set:
                messagebox.showinfo('Job Successfully finished', 'Successfully renamed blocks')
        self.frame.after(100, self.poll_log_queue)

