# Sharepoint Downloader

A script that utilizes aria2 to download sharepoint links you received.

---

### Why would I want to use this script?

1. You can batch-download hundred, or thousands of sharepoint links automatically (after setup).
2. You can continue file download when your Internet goes bad, or sometimes the download just... disconnected for no reason.
3. I once tried to download a sharepoint folder manually, and I got throttled after downloading like 2x files out of hundred files, I had to check the page every few minutes myself... and I don't know after how long, it finally makes me download again.  
   As you might guess, it happened more than few times, throttle and waiting makes me crazy.
4. You might say sharepoint has a download button when you browse folder page.... it's useful when you download a sharepoint folder under like 1GB, but if the files size goes up (like 10GBs or 200GBs), it will take time and harddrive space to extract the files you downloaded (twice if you are downloading multti-part rar file), not to mention there is no continue downloading... if you lost connection you startover.

---

### Demo

![](https://i.imgur.com/IUtATbW.gif)

---

### Requirements
1. [Node.js](https://nodejs.org) (I tested with version 10 and 12 during developement, but it might be able to run on older version.)
2. [Aria2](https://aria2.github.io/) (I tested with 1.34.0)
3. Few minutes to setup for yourself.

---

### Quick setup guide

1. [Download this project](https://github.com/axzxc1236/SharepointDownloader/archive/master.zip)
2. Open aria2.conf with a text editor, change "dir=" section and "rpc-secret=" section, save file.
3. Open aria2.conn.config with a text editor, change "Aria2_token=" section, save file.
4. Open tasklist with a text editor, you can add tasks there or don't change anything so you can test this script first.
5. Open a command terminal, do `(full path to aria2c.exe) --conf-path=(full path to aria2.conf)`, you need to change two places in that command for it to work.
6. Open another command terminal, do `cd (path to the code folder)` and then do `npm i`.
7. Do `node ./` in the seconds terminal after previous command completes.
8. There should be some texts on your second command terminal.


##### Donations?

If my work helps you so much that you want to show some appreciation... please consider donate to me.  (or just star this project)  
Writing this script is not easy... to put it lightly.  
Hours and hours have been poured into writing this script.

Any donation will motivate me.  
BTC: bc1q4fxqncgd89a7pear2mvhz4y2rmsxladmhz4pkl  
ETH: 0xd0ccb7caecf9d7ad3a3fe2042c83a2e1eba2af40
