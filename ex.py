import customtkinter as ctr


ctr.set_default_color_theme("blue")
ctr.set_appearance_mode("Native")


root = ctr.CTk()
root.geometry("1500x900+250+60")
root.title("Kinnari")
root.resizable(width=False, height=False)

dumbass = ctr.IntVar() # <------------- THIS ONE


def get_value(*args):
    print(dumbass.get()) #<------------------

settingFrame = ctr.CTkFrame(root, width=424, height=585)
settingFrame.place(x=24, y=24)

label = ctr.CTkLabel(settingFrame, text="Brightness: ", font=("", 18))
label.place(x=40, y=35)

sliderBright = ctr.CTkSlider(
    settingFrame, from_=0, to=100, variable=dumbass, command=get_value
)    #<--------------------- HERE
sliderBright.place(x=40, y=90)


root.mainloop()