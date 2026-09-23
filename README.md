# ROBOCON templates

仓库中的模板文件可以直接用 PowerPoint 或 WPS 编辑。需要通过代码批量修改时，使用 `src/customize_template.py`。

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
- `ROBOCON_清简风格模板.pptx`：清简风格演示文稿模板。
- `ROBOCON_清简风格模板.potx`：清简风格 PowerPoint 模板文件。
