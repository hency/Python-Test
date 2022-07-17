while True:
    s=input('请输入一个词：')
    if s=='退出':
     break
    if len(s)<2:
        print('输入的词太短')
        continue
    print('输入的词足够长')