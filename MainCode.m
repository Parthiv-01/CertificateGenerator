clc
clear all
close all
home

filename = 'Registration_Details.xls';
[num,txt] = xlsread(filename)

len=length(txt)

blankimage = imread('certificate.jpeg');

for i=1:len
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end

for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end

for i=1:len
    for j= 4:4
      text_date(i,j)=txt(i,j)
    end
end

for i=2:len
        text_all=[text_names(i,2) text_topic(i,3) text_date(i,4)]
        
        position = [469 380;660 480; 820 540];         
        
        RGB = insertText(blankimage,position,text_all,'FontSize',28,'BoxOpacity',0);
        
        figure
        imshow(RGB)        

        y=i-1
        filename=['Certificate_Topic_' num2str(y)]
        lastf=[filename '.png']
        saveas(gcf,lastf)
end    

