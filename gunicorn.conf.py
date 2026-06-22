import multiprocessing

# Timeout raised to 180s — GST reconciliation with large ZIP files
# (14+ PDFs parsed in parallel) can easily exceed 120s on Render free tier.
# This value must also be passed explicitly via --timeout in the start
# command in render.yaml, since Render sometimes ignores gunicorn.conf.py.
timeout = 180

# Workers: keep to 1 on free tier to stay within 512MB RAM
workers = 1

# Bind
bind = "0.0.0.0:10000"
