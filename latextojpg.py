import matplotlib.pyplot as plt
from os import mkdir
from os.path import exists
from PIL import Image, ImageChops
from matplotlib import rcParams

config = {
    "font.family":'serif',
    "font.size": 20,
    "mathtext.fontset":'stix',
    "font.serif": ['SimSun'],
}
rcParams.update(config)

def trim(im2):
    bg = Image.new(im2.mode, im2.size, im2.getpixel((0, 0)))
    diff = ImageChops.difference(im2, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im2.crop(bbox)


def latex2png(str_latex, out_file, font_size=16):
    fig = plt.figure(figsize=(20, 10), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.get_xaxis().set_visible(False)
    ax.get_yaxis().set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_xticks([])
    ax.set_yticks([])
    plt.text(0.5, 0.5, str_latex, fontsize=font_size, verticalalignment='center', horizontalalignment='center')
    plt.axis('off')
    plt.savefig(out_file)


path = 'images'
if not exists(path):
    mkdir(path)
mathtemp = r'$\alpha_v = (4.0 - 0.08d) \sqrt{\frac{f_c}{f_y}} = 0.444$'
latex2png(mathtemp, f'{path}/fwn.jpg', font_size=16)

im = Image.open(f'{path}/fwn.jpg')
im = trim(im)
im.save(f'{path}/fwn.jpg')