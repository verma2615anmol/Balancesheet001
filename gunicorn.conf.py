import multiprocessing

# Timeout: large workbooks (65K+ rows) need up to 60s to process
timeout = 120

# Workers: keep to 1 on free tier to stay within 512MB RAM
workers = 1

# Bind
bind = "0.0.0.0:10000"
