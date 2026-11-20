Print Tracker - Kyocera TASKalfa 2553ci
=======================================

START
  Double-click "Start Print Tracker.bat". A black window opens and the dashboard
  opens in your browser at http://localhost:8765
  Keep the black window open (you can minimise it). Closing it stops live tracking,
  but prints made meanwhile are still picked up from the printer counter next time.

FILES
  app.py              the tracker (reads printer counters + Windows print queue)
  dashboard.html      the dashboard and the Pricing & settings page
  config.json         created on first start: printer IP, port
  print_tracker.db    created on first start: all tracked prints and your rates

NOTES
  * Counts come from the printer's own counters at http://192.168.50.10
  * Document names and tray (bypass / cassette) come from the Windows print queue.
  * Rates are set in the dashboard under "Pricing & settings".
