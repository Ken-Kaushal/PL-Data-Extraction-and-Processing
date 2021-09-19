%%Input .txt file name for data collection
test = '20210506_165109_spectrum_all.txt';
[pathstr,name,ext] = fileparts(test);
ext = '.xlsx';
%%Input .xlsx file name for data output
s = strcat(name,ext);
filename = s;
%%Variable allocation
T = readmatrix(test);
r = size(T,1);
T(r+1,1:3) = 0;
r = size(T,1);
c = r/1044;
c = fix(c) - 1;
%%Start Data Collection and Conversion
%%For Wavelength Vs Count
J = T;
    for j = 1:c
      for i = 1:1044
          T(i+1,j+2) = J(1045*j+i,2);
      end
      T(1,j+1) = j;
    end
    T(:,1:2) = 0;
    for i = 1:1044
        for j = 1:2
            T(i+1,j) = J(i,j);
        end
    end
  T(1,1) = 0;  
  T(1,2) = 1;
  T(1,c+2) = c+1;
    
C = T(1:1045,1:c+2);
B = cast(C,'single');
C = T(2:1045,2:c+2);
Nr = normalize(C,'range');
NrC = Nr;
for j = 1:c+1
    for i = 1:1044
        Nr(i+1,j)= NrC(i,j);
    end
    Nr(1,j) = j;
end
       
%filename = 'spectrum_all_export.xlsx';
writematrix(B,filename,'Sheet','Counts','Range','A2:NY1046')
writematrix(Nr,filename,'Sheet','Normalized_Count','Range','A2:NX1046')

%%For Absolute irradiance
T = J;
T(:,1) = 0;

for j = 1:c+1
      for i = 1:1044
          T(i+1,j+1) = J(1045*(j-1)+i,3);
          T(i+1,1) = J(i,1);
      end
      T(1,j+1) = j;
end

C = T(1:1045,1:c+2);
B = cast(C,'single');
C = T(2:1045,2:c+2);
C = max(C,0);
Nr = normalize(C,'range');
NrC = Nr;
for j = 1:c+1
    for i = 1:1044
        Nr(i+1,j)= NrC(i,j);
    end
    Nr(1,j) = j;
end

%filename = 'spectrum_all_export.xlsx';
writematrix(B,filename,'Sheet','Abs_Irradiance','Range','A2:NY1046')
writematrix(Nr,filename,'Sheet','Normalized_Abs_Irradiance','Range','A2:NX1046')
%writematrix(Nr,filename,'Sheet','Normalized','Range','NZ2:ACW1045')

        
    
    

% for n = 1:387
%     for i = 1:1045
%         for j = 1:3
%             L(i,j) = T(1045*n+(i-1),j);
%         end
%     end
%     k = (4*(n-1)+5);
%     for i = 1:1045
%             for l = 1:3
%                 
%                      T(i,k+l-1)=L(i,l);                
%             end
%     end
% end
% 
%  R = T; 
%  
%  
%   for i = 1:40560
%       if i < 1046
%       for m = 1:3
%           T(i+1,m) = R(i,m);
%       end
%       else
%           for m = 1:3
%               T(i,m) = 0;
%           end
%       end
%            
%   end
%   T(1,:) = 0;
%   T(1,1) = 1;
%   
%   for n = 1:387
%       k = (4*(n-1)+5);
%       T(1,k) = n+1;
%   end
%   
%   R = T(1:1045,1:1551);
%  
%  filename = 'spectrum_all_export.xlsx';
%  writematrix(R,filename)


  
  











    
             
       

