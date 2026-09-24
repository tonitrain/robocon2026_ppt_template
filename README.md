# ROBOCON templates

仓库中的模板文件可以直接用 PowerPoint 或 WPS 编辑。需要通过代码批量修改时，使用 `src/customize_template.py`。

## 自定义模板风格

`ROBOCON_template_editable_v4.pptx` 是解锁版模板，母版和版式中的装饰元素已经放到普通幻灯片层。克隆仓库后可以直接用 PowerPoint 或 WPS 移动、复制、删除和修改数字、横线、竖线、色块、字母、图片等元素。

如果需要从原始模板重新生成解锁版：

```powershell
python src/promote_design_elements.py `
  --input ROBOCON_template_no_fixed_labels_v3.pptx `
  --output ROBOCON_template_editable_v4.pptx
```

然后按下面的方式继续自定义风格：

1. 在 PowerPoint 中打开 `ROBOCON_template_no_fixed_labels_v3.pptx`。
2. 进入“视图 -> 幻灯片母版”，修改主题字体、主题颜色、背景、线条、图片、版式和占位符。
3. 回到普通视图，修改文字、图片、图形和页面结构。
4. 使用“另存为”保存为新的 `.pptx` 或 `.potx`，再把新的模板提交到自己的分支。

母版中的元素需要在“幻灯片母版”视图中修改，普通幻灯片中的元素可以直接修改。仓库中的代码脚本是批量替换文字的辅助工具；完整的视觉风格编辑以 PPTX/POTX 文件为准。

## 代码修改

先查看模板中的文字：

```powershell
python src/customize_template.py `
  --input ROBOCON_template_no_fixed_labels_v3.pptx `
  --list-text
```

替换文字并生成新文件：

```powershell
python src/customize_template.py `
  --input ROBOCON_template_no_fixed_labels_v3.pptx `
  --output 我的项目汇报.pptx `
  --replace "[项目名称]=我的项目" `
  --replace "团队名称=我的团队" `
  --replace "汇报人=张三"
```

也可以把替换内容放进 JSON 文件：

```json
{
  "[项目名称]": "我的项目",
  "团队名称": "我的团队",
  "汇报人": "张三"
}
```

```powershell
python src/customize_template.py `
  --input ROBOCON_template_no_fixed_labels_v3.pptx `
  --output 我的项目汇报.pptx `
  --replace-file replacements.json
```

脚本会扫描幻灯片、版式和母版中的文字。输入文件始终保留，输出文件是新的 PPTX 或 POTX 文件；修改版式中的文字时，替换内容需要与文本框中的完整文字一致。

## 文件

- `ROBOCON_template_no_fixed_labels_v3.pptx`：已去除固定标识文字的汇报模板。
- `ROBOCON_template_editable_v4.pptx`：已将母版和版式装饰元素解锁到普通幻灯片层的模板。
- `ROBOCON_清简风格模板.pptx`：清简风格演示文稿模板。
- `ROBOCON_清简风格模板.potx`：清简风格 PowerPoint 模板文件。
