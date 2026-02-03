clear; clc; close all; warning off; % 环境初始化

% ===================== 1. 基础设置（可按需修改）=====================
excel_file = '123.xlsx';    % Excel文件名（确保和代码在同一目录）
sheet_name = 'Sheet1';        % 工作表名（按需修改）
save_folder = 'distribution_plots'; % 图片保存文件夹（自动创建）
dpi = 600;                    % 图片分辨率（固定600dpi）
figure_size = [100, 100, 1500, 1200]; % 增大图形尺寸，适配30号大字体
font_size = 30;               % 统一字体大小（标题/标签/图例/刻度）

% ===================== 2. 初始化：创建保存文件夹=====================
if ~exist(save_folder, 'dir')
    mkdir(save_folder);
    fprintf('已创建图片保存文件夹：%s\n', fullfile(pwd, save_folder));
else
    fprintf('图片将保存至已存在文件夹：%s\n', fullfile(pwd, save_folder));
end

% ===================== 3. 导入表格并筛选数值列=====================
data_table = readtable(excel_file, 'Sheet', sheet_name);
all_cols = data_table.Properties.VariableNames;
numeric_cols = {};

% 筛选数值列（跳过非数值列如“工艺名称”）
for i = 1:length(all_cols)
    col_data = data_table.(all_cols{i});
    if isnumeric(col_data)
        numeric_cols{end+1} = all_cols{i};
    end
end

if isempty(numeric_cols)
    error('表格中未找到数值列，请检查工作表或列数据类型！');
end
fprintf('\n共筛选出 %d 个数值列，将批量处理：\n', length(numeric_cols));
fprintf('列名列表：%s\n\n', strjoin(numeric_cols, ', '));

% ===================== 4. 批量处理每个数值列=====================
processed_count = 0;
for col = numeric_cols
    col = col{1};
    processed_count = processed_count + 1;

    % -------------------- 4.1 提取并清理数据
    data = data_table.(col);
    data = data(~isnan(data)); % 移除空值
    if length(data) < 5
        fprintf('[%d/%d] 跳过：%s列（有效数据仅%d个，需≥5个）\n', ...
                processed_count, length(numeric_cols), col, length(data));
        continue;
    end

    % -------------------- 4.2 计算统计量（命令行输出不变）
    mu = mean(data);
    sigma = std(data);
    sk = skewness(data);
    ku = kurtosis(data);
    sample_num = length(data);

    fprintf('[%d/%d] 正在处理：%s列（有效样本数：%d）\n', ...
            processed_count, length(numeric_cols), col, sample_num);
    fprintf('        均值：%.4f | 标准差：%.4f | 偏度：%.4f | 峰度：%.4f\n', ...
            mu, sigma, sk, ku);

    % -------------------- 4.3 生成图形（无英文+30号字体）
    figure('Position', figure_size); % 大尺寸适配大字体
    subplot(1,1,1); hold on; grid on;
    grid minor; % 补充细网格，提升可读性（可选保留）

    % 绘制直方图
    histogram(data, 'Normalization', 'pdf', ...
              'EdgeColor', 'none', 'FaceAlpha', 0.7, ...
              'BinWidth', range(data)/30);

    % 绘制正态分布曲线
    x_norm = linspace(min(data), max(data), 1000);
    y_norm = normpdf(x_norm, mu, sigma);
    plot(x_norm, y_norm, 'LineWidth', 4, 'Color', [0.9, 0.2, 0.2]); % 加粗曲线适配大图

    % 绘制核密度估计曲线
    [f_kde, xi_kde] = ksdensity(data);
    plot(xi_kde, f_kde, 'LineWidth', 4, 'Color', [0.1, 0.5, 0.8]); % 加粗曲线

    % -------------------- 图形文字设置（无英文+30号字体）
    % 标题：仅列名（移除“Distribution of”英文）
    title(col, 'FontSize', font_size, 'FontWeight', 'bold');
    % 坐标轴标签：纯中文+列名，30号字体
    xlabel(sprintf('%s 数值', col), 'FontSize', font_size);
    ylabel('概率密度', 'FontSize', font_size);
    % 图例：中文+30号字体，避免遮挡
    legend('直方图', '正态分布', '核密度估计', ...
           'Location', 'best', 'FontSize', font_size, 'Box', 'off', 'NumColumns', 1);
    % 坐标轴刻度：30号字体（确保刻度文字清晰）
    set(gca, 'FontSize', font_size, 'GridAlpha', 0.3, 'LineWidth', 2); % 加粗坐标轴

    % -------------------- 4.4 保存高清图片
    img_name = sprintf('%s_分布.png', col); % 文件名改为中文（可选）
    img_path = fullfile(save_folder, img_name);
    print(gcf, img_path, '-dpng', sprintf('-r%d', dpi));
    fprintf('        已保存图片：%s\n\n', img_path);

    close(gcf); % 关闭图形释放内存
end

% ===================== 5. 处理完成汇总=====================
fprintf('===================== 批量处理完成 =====================\n');
fprintf('总数值列数：%d | 成功处理列数：%d\n', length(numeric_cols), processed_count);
fprintf('所有图片保存路径：%s\n', fullfile(pwd, save_folder));
fprintf('图片格式：PNG（%ddpi） | 字体大小：%d号 | 无英文文字\n', dpi, font_size);