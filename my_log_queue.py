import logging

class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    
    It can be used from different threads
    """
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
        
        def emit(self, record):
            self.log_queue.put(record)



logger = logging.getLogger(__name__)
print(f"XXX - my_log_queue.logger:{logger}")
