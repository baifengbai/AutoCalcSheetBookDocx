import os

os.chdir(os.path.dirname(__file__))
outfile = open('dir.txt', 'a')                          # 以追加方式打开输出文件
for dirpath, dirs, files in os.walk('.'):                # 递归遍历当前目录和所有子目录的文件和目录
    for name in files:                                   # files保存的是所有的文件名
        if os.path.splitext(name)[1] == '.xls' or os.path.splitext(name)[1] == '.xlsx':        
            #filename = os.path.join(dirpath, name)       # 加上路径，dirpath是遍历时文件对应的路径
            #f = open(filename, 'r')
            print(name)
            outfile.write(name + '\n')                          # 写入输出文件
            #f.close()    
outfile.close()