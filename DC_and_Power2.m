% Import data from Excel file
A = importdata('C:\Temp\SMU_LI\20230928_152011_EwokMZPLUS_CornerSample Sample2_3.3V_15C_2ns.xlsx');
Data = A.data;

% Extract columns to be used
Selion = Data(:,1);
mP = Data(:,2);

% Define duty cycle values for interpolation
T = 24;
DCm = [10 20 30 40 50 63; 1.27490/T 1.64877/T 1.77782/T 1.82697/T 1.91134/T 1.9059/T];
DCm = DCm';
xq = 0:1:63;

% Interpolate duty cycle values
DC = interp1(DCm(:,1), DCm(:,2), xq);
DC1 = pchip(DCm(:,1), DCm(:,2), xq);
DC2 = spline(DCm(:,1), DCm(:,2), xq);

% Calculate peak power
P = mP./DC1';

% Plot duty cycle and peak power
figure
subplot(2,1,1)
plot(xq, DC, 'b+', xq, DC1, 'r*', xq, DC2, 'gx', DCm(:,1), DCm(:,2), '.')
ylabel('Duty Cycle')
xlabel('Selion')
legend('Interpolated DC', 'PCHIP', 'Spline', 'Measured DC')

subplot(2,1,2)
plot(Selion, P, Selion, Data(:,7))
ylabel('Peak Power')
xlabel('Selion')
legend('Interpolated P', 'Measured P')
