%penyelesaian kasus menggunakan metode WP
clc;clear; %untuk membersihkan jendela command windows

opts = detectImportOptions('wp_123190133.xlsx'); %mengimport data excel
opts.SelectedVariableNames = (3:5);%mengambil data dari kolom 3-5
data1 = readmatrix('wp_123190133.xlsx',opts); %membaca matrix dari data

opts = detectImportOptions('wp_123190133.xlsx'); %mengimport data excel
opts.SelectedVariableNames = (8); %mengambil data dari kolom 8
data2 = readmatrix('wp_123190133.xlsx',opts); %membaca matrix dari data

data = [data1 data2];
data = data(1:50,:); %mengambil data dari no 1 sampai 50

x = [data]; %data rating kecocokan dari masing-masing alternatif
k = [0,0,1,0]; %atribut tiap-tiap kriteria, dimana nilai 1=atrribut keuntungan, dan 0= atribut biaya
w = [3,5,4,1]; %Nilai bobot tiap kriteria (1= sangat buruk, 2=buruk, 3=cukup, 4= tinggi, 5= sangat tinggi)

%tahapan pertama, perbaikan bobot
[m n]=size(x); %inisialisasi ukuran x
w=w./sum(w); %membagi bobot per kriteria dengan jumlah total seluruh bobot

%tahapan kedua, melakukan perhitungan vektor(S) per baris (alternatif)
for j=1:n,
    if k(j)==0, w(j)=-1*w(j);
    end;
end;
for i=1:m,
    S(i)=prod(x(i,:).^w);
end;

%tahapan ketiga, proses perangkingan
V= S/sum(S)

%tahapan keempat, mengubah menjadi v transpose
Vtranspose = V.';
opts = detectImportOptions('wp_123190133.xlsx'); %mengimport data excel
opts.SelectedVariableNames = (1);%mengambil data dari kolom 1
data3 = readmatrix('wp_123190133.xlsx',opts); %membaca matrix dari data
data3 = data3(1:50,:); %mengambil data dari no 1 sampai 50
data4 = [data3 Vtranspose];
data4 = sortrows(data4,-2);
data4 = data4(1:5,1);%mengambil data dari rangking 1-5

disp ('Nomer Real Estate yang menjadi alternatif terbaik :')
disp (data4)