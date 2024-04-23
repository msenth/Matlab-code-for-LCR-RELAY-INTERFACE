function lcr_matlab_8freq
delete(instrfind({'Port'},{'COM4'}));
%am=serial('COM4','BaudRate',9600); %Creating the serial port object
%fopen(am); %Opening the port
%resourcelist=visadevlist; % Get all device addreses
lcr=visadev("USB0::0x0D4A::0x003F::9300676::0::INSTR"); % Use address of LCR here
%freqlist=[100 1000 10000 100000 1000000]; 
freqlist=[100 1000 2000 5000 10000 20000 100000 1000000]; 
%freqlist=[100 200 500 1000 2000 5000 10000 20000 50000 100000 200000 500000 1000000]; 
formatspec=":SOUR:FREQ %d";
volt = 0.1;
formatspec1 = ":SOUR:VOLT %d";
runFor = 24;
T=zeros(1,8);
zmag=zeros(1,8);
theta=zeros(1,8);
readstr=string();
for i = 1:4
    clc;
   % fwrite(am, [255 i 1]);
    fname=sprintf('LF_healthy3__12_01_23_%01d.xlsx',i);
    %fname=sprintf('05_09_23_oncho_selectindivsamples_30mins_%01d.xlsx',i);
%% Measurement
    tic;
    figure;
    count = zeros(1,8);
    
    global zmag_cellarr;
    time_cellarr = {0};
    zmag_cellarr = {0};
    hold on;
    while (toc < runFor)
        for k = 1:length(freqlist)
         sendstr=sprintf(formatspec, freqlist(k));
            sendstr1=sprintf(formatspec1,volt);
            write(lcr,sendstr); % set measurement frequency
            write(lcr,sendstr1); %set measurement voltage
            writeread(lcr,"*TRG"); % trigger and get reading
            readstr(k)=ans; 
             %reading from lcr
            readnum=sscanf(readstr(k),'%f,%f,%f');
            zmag(k)=readnum(2);
            theta(k)=readnum(3);
 
            T(k) = toc; %pair 1: toc - stores timestamp in array
 
        %impedance vs time
            title('Impedance Measurement vs. Time');
            xlabel('Time (s)');
            ylabel('Impedance (Ohms)'); %change labels depending on what is being measured
             if freqlist(k) == freqlist(1)
                if count(1) == 0
                    zmag_cellarr{1} = zmag(k);
                    time_cellarr{1} = T(k);
                    theta_cellarr{1} = theta(k);
                end
                count(1) = count(1) + 1;
                zmag_cellarr{1} = [zmag_cellarr{1}, zmag(k)];
                time_cellarr{1} = [time_cellarr{1}, T(k)];
                theta_cellarr{1} = [theta_cellarr{1}, theta(k)];
                if count(1) == 1
                    
                    %data_table = {'Time',sprintf('%d Hz', freqlist(1)), sprintf('%d Hz', freqlist(2)), sprintf('%d Hz', freqlist(3)), sprintf('%d Hz', freqlist(4)),sprintf('%d Hz', freqlist(5))};
                    data_table = {'Time',sprintf('%d Hz', freqlist(1)), sprintf('%d Hz', freqlist(2)), sprintf('%d Hz', freqlist(3)), sprintf('%d Hz', freqlist(4)),sprintf('%d Hz', freqlist(5)),sprintf('%d Hz', freqlist(6)), sprintf('%d Hz', freqlist(7)), sprintf('%d Hz', freqlist(8))};
                    %data_table = {'Time',sprintf('%d Hz', freqlist(1)), sprintf('%d Hz', freqlist(2)), sprintf('%d Hz', freqlist(3)), sprintf('%d Hz', freqlist(4)),sprintf('%d Hz', freqlist(5)),sprintf('%d Hz', freqlist(6)), sprintf('%d Hz', freqlist(7)), sprintf('%d Hz', freqlist(8)), sprintf('%d Hz', freqlist(9)), sprintf('%d Hz', freqlist(10)), sprintf('%d Hz', freqlist(11)), sprintf('%d Hz', freqlist(12)), sprintf('%d Hz', freqlist(13))};
                    %writecell(data_table,'fname','Sheet',1,'Range','A1');
                    writecell(data_table,fname,'Sheet',1,'Range','A1');
                end
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{1}]','Sheet1','B2');
                    xlswrite(fname,[theta_cellarr{1}]','Sheet1','Q2');
                p1 = loglog(time_cellarr{1}(1:length(time_cellarr{1})),zmag_cellarr{1}(1:length(zmag_cellarr{1})));
            end
            if freqlist(k) == freqlist(2)
                if count(2) == 0
                    zmag_cellarr{2} = zmag(k);
                    time_cellarr{2} = T(k);
                    theta_cellarr{2} = theta(k);
                end
                count(2) = count(2)+1;
                zmag_cellarr{2} = [zmag_cellarr{2}, zmag(k)];
                time_cellarr{2} = [time_cellarr{2}, T(k)];
                theta_cellarr{2} = [theta_cellarr{1}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{2}]','Sheet1','C2');
                    xlswrite(fname,[theta_cellarr{2}]','Sheet1','R2');
 
                p2 = loglog(time_cellarr{2}(1:length(time_cellarr{2})), zmag_cellarr{1}(1:length(zmag_cellarr{2})));
            end 

            if freqlist(k) == freqlist(3)
                if count(3) == 0
                    zmag_cellarr{3} = zmag(k);
                    time_cellarr{3} = T(k);
                    theta_cellarr{1} = theta(k);
                end
                count(3) = count(3) + 1;
                zmag_cellarr{3} = [zmag_cellarr{3}, zmag(k)];
                time_cellarr{3} = [time_cellarr{3}, T(k)];
                theta_cellarr{3} = [theta_cellarr{1}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{3}]','Sheet1','D2');
                    xlswrite(fname,[theta_cellarr{3}]','Sheet1','S2');
                p3 = loglog(time_cellarr{3}(1:length(time_cellarr{3})), zmag_cellarr{3}(1:length(zmag_cellarr{3})));
            end 
 
            if freqlist(k) == freqlist(4)
                if count(4) == 0
                    zmag_cellarr{4} = zmag(k);
                    time_cellarr{4} = T(k);
                    theta_cellarr{4} = theta(k);
                end
                count(4) = count(4) + 1;
                zmag_cellarr{4} = [zmag_cellarr{4}, zmag(k)];
                time_cellarr{4} = [time_cellarr{4}, T(k)];
                theta_cellarr{4} = [theta_cellarr{4}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{4}]','Sheet1','E2');
                    xlswrite(fname,[theta_cellarr{4}]','Sheet1','T2');
                p4 = loglog(time_cellarr{4}(1:length(time_cellarr{4})), zmag_cellarr{4}(1:length(zmag_cellarr{4})));
            end 

            if freqlist(k) == freqlist(5)
                if count(5) == 0
                    0
                    zmag_cellarr{5} = zmag(k);
                    time_cellarr{5} = T(k);
                    theta_cellarr{5} = theta(k);
                end
                count(5) = count(5) + 1;
                zmag_cellarr{5} = [zmag_cellarr{5}, zmag(k)];
                time_cellarr{5} = [time_cellarr{5}, T(k)];
                theta_cellarr{5} = [theta_cellarr{5}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{5}]','Sheet1','F2');
                    xlswrite(fname,[theta_cellarr{5}]','Sheet1','U2');
                p5 = loglog(time_cellarr{5}(1:length(time_cellarr{5})), zmag_cellarr{5}(1:length(zmag_cellarr{5})));
            end 

            if freqlist(k) == freqlist(6)
                if count(6) == 0
                    zmag_cellarr{6} = zmag(k);
                    time_cellarr{6} = T(k);
                    theta_cellarr{6} = theta(k);
                end
                count(6) = count(6) + 1;
                zmag_cellarr{6} = [zmag_cellarr{6}, zmag(k)];
                time_cellarr{6} = [time_cellarr{6}, T(k)];
                theta_cellarr{6} = [theta_cellarr{6}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{6}]','Sheet1','G2');
                    xlswrite(fname,[theta_cellarr{6}]','Sheet1','V2');
                p6 = loglog(time_cellarr{6}(1:length(time_cellarr{6})), zmag_cellarr{6}(1:length(zmag_cellarr{6})));
            end 

            if freqlist(k) == freqlist(7)
                if count(7) == 0
                    zmag_cellarr{7} = zmag(k);
                    time_cellarr{7} = T(k);
                    theta_cellarr{7} = theta(k);
                end
                count(7) = count(7) + 1;
                zmag_cellarr{7} = [zmag_cellarr{7}, zmag(k)];
                time_cellarr{7} = [time_cellarr{7}, T(k)];
                theta_cellarr{7} = [theta_cellarr{7}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{7}]','Sheet1','H2');
                    xlswrite(fname,[theta_cellarr{7}]','Sheet1','W2');
                p7 = loglog(time_cellarr{7}(1:length(time_cellarr{7})), zmag_cellarr{7}(1:length(zmag_cellarr{7})));
            end 

            if freqlist(k) == freqlist(8)
                if count(8) == 0
                    zmag_cellarr{8} = zmag(k);
                    time_cellarr{8} = T(k);
                    theta_cellarr{8} = theta(k);
                end
                count(8) = count(8) + 1;
                zmag_cellarr{8} = [zmag_cellarr{8}, zmag(k)];
                time_cellarr{8} = [time_cellarr{8}, T(k)];
                theta_cellarr{8} = [theta_cellarr{8}, theta(k)];
                    xlswrite(fname,[time_cellarr{1}]','Sheet1','A2');
                    xlswrite(fname,[zmag_cellarr{8}]','Sheet1','I2');
                    xlswrite(fname,[theta_cellarr{8}]','Sheet1','X2');
                p8 = loglog(time_cellarr{8}(1:length(time_cellarr{8})), zmag_cellarr{8}(1:length(zmag_cellarr{8})));
            end
        end
    end
  %  fwrite(am, [255 i 0]);
    pause(10);
end
 %% Set up Legend
%Legend
legend([p1 p2 p3 p4 p5 p6 p7 p8], sprintf('%d Hz', freqlist(1)), sprintf('%d Hz', freqlist(2)), sprintf('%d Hz', freqlist(3)), sprintf('%d Hz', freqlist(4)),sprintf('%d Hz', freqlist(5)),sprintf('%d Hz', freqlist(6)), sprintf('%d Hz', freqlist(7)), sprintf('%d Hz', freqlist(8)));
%legend([p1 p2 p3 p4 p5], sprintf('%d Hz', freqlist(1)), sprintf('%d Hz', freqlist(2)), sprintf('%d Hz', freqlist(3)), sprintf('%d Hz', freqlist(4)),sprintf('%d Hz', freqlist(5)));