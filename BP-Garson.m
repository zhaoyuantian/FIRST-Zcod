%%  清空环境变量
warning off             % 关闭报警信息
close all               % 关闭开启的图窗
clear                   % 清空变量
clc                     % 清空命令行
%%  导入数据
data = xlsread('贡献系数.xlsx','sheet1');
length_res = size(data, 1);
wide_res = size(data, 2);

%%  划分训练集和测试集
% 打乱数据集
temp = randperm(length_res);
% 划分训练集，P_train不包括输出值，T_train为输出值，0.7表示70%数据划分为训练集
P_train = data(temp(1: 0.7*length_res), 1:wide_res-1)';
T_train = data(temp(1: 0.7*length_res), wide_res)';
M = size(P_train, 2);
% 划分测试集
P_test = data(temp(0.7*length_res + 1:end), 1:wide_res-1)';
T_test = data(temp(0.7*length_res + 1:end), wide_res)';
N = size(P_test, 2);

%%  数据归一化
[p_train, ps_input] = mapminmax(P_train, 0, 1);
p_test = mapminmax('apply', P_test, ps_input);
[t_train, ps_output] = mapminmax(T_train, 0, 1);
t_test = mapminmax('apply', T_test, ps_output);

%%  创建网络
% 5为神经单元数
net = newff(p_train, t_train, 5);

%%  设置训练参数
net.trainParam.epochs = 1000;     % 迭代次数 
net.trainParam.goal = 1e-6;       % 误差阈值
net.trainParam.lr = 0.01;         % 学习率

%%  训练网络
net = train(net, p_train, t_train);

%%  仿真测试
t_sim1 = sim(net, p_train);
t_sim2 = sim(net, p_test);

%%  数据反归一化
T_sim1 = mapminmax('reverse', t_sim1, ps_output);
T_sim2 = mapminmax('reverse', t_sim2, ps_output);

%%  均方根误差
error1 = sqrt(sum((T_sim1 - T_train).^2) ./ M);
error2 = sqrt(sum((T_sim2 - T_test).^2) ./ N);

%%  绘图
figure
plot(1:M, T_train, 'r-*', 1:M, T_sim1, 'b-o', 'LineWidth', 1)
legend('真实值', '预测值')
xlabel('预测样本')
ylabel('预测结果')
string = {strcat('训练集预测结果对比：', ['RMSE=' num2str(error1)])};
title(string)
xlim([1, M])
grid

figure
plot(1:N, T_test, 'r-*', 1:N, T_sim2, 'b-o', 'LineWidth', 1)
legend('真实值', '预测值')
xlabel('预测样本')
ylabel('预测结果')
string = {strcat('测试集预测结果对比：', ['RMSE=' num2str(error2)])};
title(string)
xlim([1, N])
grid

%%  相关指标计算
%  R2
R1 = 1 - norm(T_train - T_sim1)^2 / norm(T_train - mean(T_train))^2;
R2 = 1 - norm(T_test - T_sim2)^2 / norm(T_test - mean(T_test))^2;

disp(['训练集数据的R2为：', num2str(R1)]);
disp(['测试集数据的R2为：', num2str(R2)]);

%  MAE
mae1 = sum(abs(T_sim1 - T_train)) ./ M;
mae2 = sum(abs(T_sim2 - T_test)) ./ N;

disp(['训练集数据的MAE为：', num2str(mae1)]);
disp(['测试集数据的MAE为：', num2str(mae2)]);

%  MBE
mbe1 = sum(T_sim1 - T_train) ./ M;
mbe2 = sum(T_sim2 - T_test) ./ N;

disp(['训练集数据的MBE为：', num2str(mbe1)]);
disp(['测试集数据的MBE为：', num2str(mbe2)]);

%  RMSE
disp(['训练集数据的RMSE为：', num2str(error1)]);
disp(['测试集数据的RMSE为：', num2str(error2)]);
%%  绘制散点图
sz = 25;
c = 'b';

figure
scatter(T_train, T_sim1, sz, c)
hold on
plot(xlim, ylim, '--k')
xlabel('训练集真实值');
ylabel('训练集预测值');
xlim([min(T_train) max(T_train)])
ylim([min(T_sim1) max(T_sim1)])
title('训练集预测值 vs. 训练集真实值')

figure
scatter(T_test, T_sim2, sz, c)
hold on
plot(xlim, ylim, '--k')
xlabel('测试集真实值');
ylabel('测试集预测值');
xlim([min(T_test) max(T_test)])
ylim([min(T_sim2) max(T_sim2)])
title('测试集预测值 vs. 测试集真实值')

% 误差测试
n_mis = size(T_test, 2);
for i = 1:n_mis
    mis(i) = (T_test(i) - T_sim2(i)) / T_test(i);
end
ymis1 = find(mis >= -0.1 & mis <= 0.1);
ymis2 = find(mis >= -0.05 & mis <= 0.05);
n_ymis = size(ymis1, 2);
%% 计算输入变量对输出变量的贡献系数
% 获取输入层到隐藏层的权重矩阵
input_to_hidden_weights = net.IW{1};

% 获取隐藏层到输出层的权重矩阵
hidden_to_output_weights = net.LW{2,1};

% Garson算法：计算绝对值的连接权重，并归一化
% 转置 hidden_to_output_weights 以便于计算
contribution = abs(input_to_hidden_weights)' * abs(hidden_to_output_weights)';

% 对每一列求和得到每个输入变量的总贡献
total_contribution = sum(contribution, 2);

% 归一化贡献系数
contribution_normalized = total_contribution / sum(total_contribution);

% 显示贡献系数
disp('输入变量对输出变量的贡献系数：');
for i = 1:wide_res-1
    fprintf('变量 %d 的贡献系数为: %.4f\n', i, contribution_normalized(i));
end

% 如果需要可以绘制贡献系数图
figure
bar(contribution_normalized)
set(gca, 'XTickLabel', {'变量1', '变量2', '变量3'})
xlabel('输入变量')
ylabel('贡献系数')
title('输入变量对输出变量的贡献系数')
grid on