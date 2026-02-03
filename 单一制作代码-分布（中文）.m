clear; clc; close all; warning off; % 环境初始化

% ========== 基础统计量计算 ==========
% 定义要分析的列名（只需改这里，标题自动同步）
col_name = 'Rir';  % 可替换为'UV254'/'Rr'/'Rir'等任意列名

% 导入Excel数据（Sheet1=工作表名，按需修改）
data_table = readtable('model.xlsx', 'Sheet', 'Sheet1');
data = data_table.(col_name);  % 按列名变量提取数据

mu = mean(data);
sigma = std(data);
sk = skewness(data);  % 偏度
ku = kurtosis(data);  % 峰度

% 统计结果输出（显示当前分析列名）
fprintf('==== 数据统计量（%s列）====\n', col_name);
fprintf('样本数: %d\n', length(data));
fprintf('均值: %.4f\n', mu);
fprintf('标准差: %.4f\n', sigma);
fprintf('偏度: %.4f (正态分布≈0)\n', sk);
fprintf('峰度: %.4f (正态分布≈3)\n', ku);

% ========== 图形绘制（修正分辨率设置，删除Q-Q图） ==========
% 图形尺寸：1000×800（移除无效的Resolution属性）
figure('Position', [100, 100, 1000, 800]);

% 单个子图（占满整个窗口）
subplot(1,1,1);
hold on;

% 绘制直方图
histogram(data, 'Normalization', 'pdf', ...
          'EdgeColor', 'none', 'FaceAlpha', 0.7, ...
          'BinWidth', range(data)/30);
      
% 绘制理论正态分布曲线
x = linspace(min(data), max(data), 1000);
y = normpdf(x, mu, sigma);
plot(x, y, 'LineWidth', 2, 'Color', [0.9, 0.2, 0.2]);

% 绘制核密度估计曲线
[f, xi] = ksdensity(data);
plot(xi, f, 'LineWidth', 2, 'Color', [0.1, 0.5, 0.8]);

% 图形修饰
grid on;
title(sprintf('%s', col_name), 'FontSize', 24);
xlabel(sprintf('%s 数值', col_name), 'FontSize', 30);
ylabel('概率密度', 'FontSize', 30);
legend('直方图', '正态分布', '核密度估计', 'Location', 'best', 'FontSize', 32);
set(gca, 'FontSize', 30);

% ========== 正确设置高清分辨率（600dpi）：通过print函数保存 ==========
% 保存为PNG格式，分辨率600dpi，文件名自动包含列名（避免覆盖）
save_path = sprintf('%s_distribution.png', col_name);  % 文件名示例：MF_distribution.png
print(gcf, save_path, '-dpng', '-r600');  % -r600指定分辨率600dpi
fprintf('高清图片已保存至：%s\n', save_path);