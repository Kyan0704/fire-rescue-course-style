# PPT生成脚本使用说明

## 脚本位置

`scripts/generate_ppt.py`

## 功能

根据JSON配置文件，自动生成符合消防救援课程风格的PPT文件。支持所有8种标准页面模板。

## 使用方法

```bash
python scripts/generate_ppt.py --input content.json --output output.pptx
```

### 参数

- `--input` / `-i`：输入JSON配置文件路径（必填）
- `--output` / `-o`：输出PPTX文件路径（必填）

## JSON配置格式

```json
{
  "title": "课程名称",
  "slides": [
    {
      "type": "cover",
      "data": {
        "title": "主标题",
        "subtitle": "副标题",
        "speaker": "演讲者：XXX",
        "background_image": "path/to/bg.jpg"
      }
    },
    {
      "type": "contents",
      "data": {
        "items": [
          "一、章节一名称",
          "二、章节二名称",
          "三、章节三名称"
        ]
      }
    },
    {
      "type": "chapter",
      "data": {
        "label": "CHAPTER 01 / 绪论",
        "title": "章节大标题",
        "quote": "引言文字"
      }
    },
    {
      "type": "image_text",
      "data": {
        "page_title": "页面标题",
        "image": "path/to/image.jpg",
        "card_title": "卡片标题",
        "card_body": "卡片正文内容",
        "points": [
          {"title": "要点1", "desc": "描述1"},
          {"title": "要点2", "desc": "描述2"}
        ]
      }
    },
    {
      "type": "grid_2x2",
      "data": {
        "page_title": "页面标题",
        "cards": [
          {"title": "标题1", "body": "正文1", "icon": "path/icon1.png"},
          {"title": "标题2", "body": "正文2", "icon": "path/icon2.png"},
          {"title": "标题3", "body": "正文3", "icon": "path/icon3.png"},
          {"title": "标题4", "body": "正文4", "icon": "path/icon4.png"}
        ]
      }
    },
    {
      "type": "timeline_4col",
      "data": {
        "page_title": "页面标题",
        "columns": [
          {"title": "阶段1", "time": "2010年前", "body": "描述文字"},
          {"title": "阶段2", "time": "2010-2015", "body": "描述文字"},
          {"title": "阶段3", "time": "2015-2020", "body": "描述文字"},
          {"title": "阶段4", "time": "2020至今", "body": "描述文字"}
        ]
      }
    },
    {
      "type": "two_column",
      "data": {
        "page_title": "页面标题",
        "left": {
          "title": "左栏标题",
          "image": "path/left.jpg",
          "points": [{"desc": "要点描述"}]
        },
        "right": {
          "title": "右栏标题",
          "image": "path/right.jpg",
          "points": [{"desc": "要点描述"}]
        }
      }
    },
    {
      "type": "summary",
      "data": {
        "title": "总结大标题",
        "points": [
          {"title": "要点1标题", "body": "要点1正文"},
          {"title": "要点2标题", "body": "要点2正文"}
        ],
        "thanks": "感谢聆听！"
      }
    },
    {
      "type": "end",
      "data": {
        "course_name": "课程名称",
        "speaker": "演讲者：XXX",
        "slogan": "标语文字"
      }
    }
  ]
}
```

## 页面类型说明

| type值 | 对应模板 | 必填字段 |
|--------|----------|----------|
| `cover` | 封面页 | title, subtitle, speaker |
| `contents` | 目录页 | items (数组) |
| `chapter` | 章节过渡页 | label, title, quote |
| `image_text` | 左图右文页 | page_title, image, card_title, card_body, points |
| `grid_2x2` | 四宫格页 | page_title, cards (4个) |
| `timeline_4col` | 四列时间线 | page_title, columns (4个) |
| `two_column` | 双栏对比页 | page_title, left, right |
| `summary` | 总结页 | title, points, thanks |
| `end` | 结束页 | course_name, speaker, slogan |

## 完整示例

参考 `examples/sample_content.json`（如存在）获取完整配置示例。

## 注意事项

1. 图片路径支持绝对路径和相对路径（相对路径相对于JSON文件所在目录）
2. 如不提供背景图片，将使用纯色背景
3. 图标为可选字段，不提供时使用默认圆形占位
4. 所有文字会自动应用对应风格的字体、字号、颜色
5. 生成的PPT可在PowerPoint/WPS中进一步编辑微调
