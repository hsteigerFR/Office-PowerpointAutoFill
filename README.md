# Office - Powerpoint Auto Fill

Author : Hugo Steiger

The goal of this project is to automate the creation of Powerpoint presentations that have a specific template. This tool was created and used for Mines Nancy N18's graduation ceremony. Indeed, customized slides had to be prepared for the 200 students of the promotion. The automation strategy is quite simple : 
- Create a slide template on Powerpoint
- Create an Excel Sheet where each column corresponds to a template element, and whose rows correspond to the different slides
- Iteratively copy the slide template and replace its content with the expected informations of the given row

This process is summed up in the following illustration. This graph also explains the example used in this repositery :  

![PowerpointAutoFill](https://user-images.githubusercontent.com/106969232/182206869-92a607f2-dc9c-47ff-809d-a961c9947abc.JPG)

HOW TO USE :
- Open the Powerpoint file and enable Macros
- Go to the "View" menu on the upper ribbon
- Select "Macros" on the left and choose the "Automate" macro 
- The VBA code can be checked with the "Edit" button, it is also saved as "macro.bas" in this repo
- Press Run : the slide creation process should start
