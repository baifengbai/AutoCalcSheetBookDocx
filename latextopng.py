import matplotlib.pyplot as plt
from os import mkdir
from os.path import exists


def latex2png(str_latex, out_file, img_size=(5,3), font_size=16):
    fig = plt.figure(figsize=img_size, dpi=300)
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
    plt.savefig(out_file)

path = 'images'
if not exists(path):
    mkdir(path)
mathtemp = r'$F_{WN} = p_n(h) \times C \times A$'
latex2png(mathtemp, f'{path}/fwn.png', img_size=(5, 3), font_size=16)

fig, ax = plt.subplots()
im = plt.imread(f'{path}/fwn.png')
ax.imshow(im, aspect='equal')
plt.axis('off')
# 去除图像周围的白边
height, width, channels = im.shape
# 如果dpi=300，那么图像大小=height*width
fig.set_size_inches(width / 100.0 / 3.0, height / 100.0 / 3.0)
plt.gca().xaxis.set_major_locator(plt.NullLocator())
plt.gca().yaxis.set_major_locator(plt.NullLocator())
plt.subplots_adjust(top=1, bottom=0, left=0, right=1, hspace=0, wspace=0)
plt.margins(0, 0)
plt.savefig(f'{path}/fwn.png', dpi=300)