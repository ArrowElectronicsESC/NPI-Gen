from PIL import Image

f_name = "0";
dpi = 96;

im = Image.open("{}.jpg".format(f_name));
im.save("{}_{}-dpi.jpg".format(f_name, dpi), dpi=(dpi,dpi))

def func_I_change_dpi(f_name, dpi):
    f_name_fmt = f_name[f_name.find(".")+1:len(f_name)];
    f_name_only = f_name[0:f_name.find(".")];
    
    im = Image.open("{}.{}".format(f_name_only, f_name_fmt));
    im.save("{}_{}-dpi.jpg".format(f_name_only, dpi), dpi=(dpi,dpi))
