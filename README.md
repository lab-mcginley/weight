# Behavior_Weigh animal
## Using the code
- Double click **run_weight** shortcut on the desktop.
- Deprive day default on Sunday, change it prior to running the code by manually key in 'D' to the deprive column of desired date.\
Before you enter for any other days, deprive date need to have weight first.
- For entering mouse ID, letter case does not matter, the code will force them to be upper case.

## Input from scale
- Press print botton on the scale to send weight to PC.
- If for some reason the port altered (e.g. from COM3 to COM5), change the port value at line 194:     
```ser = serial.Serial(port='COM3', timeout=1, xonxoff=True)```
- If you do not wish to use the scale input, enter ```MANUAL``` for mouse ID to manually entering weight.
If there's any erro relate to the scale, switch to manual before input any mouse ID.

## When done
- Excel file need to be closed or the result cannot be saved.
- Enter ```STOP``` to terminate the code.
  
