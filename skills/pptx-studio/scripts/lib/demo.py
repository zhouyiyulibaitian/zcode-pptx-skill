"""A sample deck that exercises every layout. Doubles as the smoke-test fixture.

Kept free of renderer imports on purpose so it can be built and validated
without a browser or an export target.
"""

from __future__ import annotations

from . import model as M


def demo_deck(theme_id: str = "ink", title: str = "示例演示") -> dict:
    charts = {
        "chart": "column", "categories": ["2021", "2022", "2023", "2024", "2025"],
        "series": [{"name": "文献量", "values": [12, 19, 27, 41, 58]}],
        "legend": False, "dataLabels": True, "grid": True, "valueFormat": "0",
        "title": "图 1  近五年相关文献数量",
    }
    return M.normalize({
        "schemaVersion": 1,
        "deck": {"title": title, "theme": theme_id, "language": "zh",
                 "size": {"w": 13.333, "h": 7.5, "preset": "16:9"},
                 "meta": {"author": ""}},
        "slides": [
            {"id": "s1", "layout": "cover", "notes": "开场：一句话说明选题与价值。",
             "content": {"tag": "毕业设计", "title": "江南古典园林空间叙事研究",
                         "subtitle": "以《景观叙事：解码海派园林》为理论基础的空间叙事转译方法",
                         "presenter": "答辩人：张三", "org": "工业设计", "date": "2026 年 6 月"}},
            {"id": "s2", "layout": "toc",
             "content": {"title": "目录", "items": [
                 {"text": "研究背景与问题", "note": "为什么研究叙事"},
                 {"text": "理论基础", "note": "景观叙事的三层结构"},
                 {"text": "研究方法", "note": "文本—空间对照分析"},
                 {"text": "案例转译", "note": "三个园子的叙事脚本"},
                 {"text": "设计验证", "note": "图纸与模型"},
                 {"text": "结论与不足"}]}},
            {"id": "s3", "layout": "bullets", "notes": "背景要短，重点在问题。",
             "content": {"title": "研究背景", "subtitle": "从“看景”到“读景”",
                         "bullets": [
                             {"text": "古典园林的当代困境：空间被保留，叙事被切断",
                              "sub": ["游客平均停留 42 分钟，动线集中于中部景区"]},
                             {"text": "既有研究偏重史料与形制，缺少可操作的空间叙事方法"},
                             {"text": "数字技术提供了新的讲述可能，但常流于效果图展示"}]}},
            {"id": "s4", "layout": "cards",
             "content": {"title": "理论基础", "columns": 3, "cards": [
                 {"tag": "01", "title": "叙事三要素", "body": "叙述者、叙述行为、接受者共同构成意义的传递链条。"},
                 {"tag": "02", "title": "空间句法", "body": "用整合度与选择度量化空间的可达性与可视性。"},
                 {"tag": "03", "title": "转译模型", "body": "将文本叙事结构映射为可建造的空间序列。"}]}},
            {"id": "s5", "layout": "chart",
             "content": {"title": "研究现状", "chart": charts,
                         "insight": "近五年相关研究增速明显，但“叙事+空间”的交叉成果仅占 8%。",
                         "bullets": [{"text": "文献集中在园林史与造园技法"},
                                     {"text": "叙事学成果多见于文学与影视领域"},
                                     {"text": "交叉研究尚缺可操作的设计路径"}],
                         "caption": "数据来源：CNKI 主题检索（2021—2025）"}},
            {"id": "s6", "layout": "image-text", "notes": "强调现场调研的样本量。",
             "content": {"title": "案例选取", "side": "left", "image": "", "image_ratio": 0.52,
                         "body": "选取三个具有典型叙事结构的江南园林作为样本，覆盖私家园林的三种规模。",
                         "bullets": [{"text": "拙政园：以水为轴的连续叙事"},
                                     {"text": "留园：收放交替的段落式叙事"},
                                     {"text": "网师园：小中见大的嵌套叙事"}],
                         "caption": "图片：现场测绘与影像记录"}},
            {"id": "s7", "layout": "kpi",
             "content": {"title": "调研数据", "metrics": [
                 {"value": "3", "unit": "座", "label": "园林样本", "delta": "覆盖三种规模"},
                 {"value": "126", "unit": "处", "label": "空间节点", "delta": "逐点测绘"},
                 {"value": "42", "unit": "位", "label": "受访者", "delta": "游客与专家"},
                 {"value": "8.4", "unit": "分", "label": "叙事清晰度", "delta": "满分 10"}]}},
            {"id": "s8", "layout": "timeline",
             "content": {"title": "研究路径", "steps": [
                 {"label": "第 1 阶段", "title": "文本分析", "text": "提取叙事结构"},
                 {"label": "第 2 阶段", "title": "空间转译", "text": "建立映射规则"},
                 {"label": "第 3 阶段", "title": "设计生成", "text": "生成空间序列"},
                 {"label": "第 4 阶段", "title": "验证迭代", "text": "使用者测试"}]}},
            {"id": "s9", "layout": "table",
             "content": {"title": "转译规则", "table": {"header": True, "rows": [
                 [{"text": "叙事要素"}, {"text": "空间对应"}, {"text": "设计手段"}],
                 [{"text": "起"}, {"text": "入口引导段"}, {"text": "收窄视线、压低视高"}],
                 [{"text": "承"}, {"text": "连续游廊"}, {"text": "线性动线与节奏重复"}],
                 [{"text": "转"}, {"text": "转折节点"}, {"text": "对景、框景与视线反转"}],
                 [{"text": "合"}, {"text": "主体水面"}, {"text": "放开界面、全景展开"}]]},
                 "note": "规则经 3 种规模验证，可按用地条件伸缩。"}},
            {"id": "s10", "layout": "compare",
             "content": {"title": "方案对比", "left": {"title": "原状", "items": [
                 "动线单一，集中在主景区", "叙事信息依靠标识牌", "节点之间缺少过渡"]},
                 "right": {"title": "转译后", "items": [
                     "四级叙事动线，分流明显", "空间自身承担叙事功能", "节点之间形成节奏过渡"]}}},
            {"id": "s11", "layout": "quote",
             "content": {"quote": "园林不是被观看的对象，而是一段被走过的句子。",
                         "author": "本研究核心主张", "source": "第三章"}},
            {"id": "s12", "layout": "big-number",
             "content": {"title": "设计验证", "value": "+37", "unit": "%",
                         "caption": "受访者对空间序列的理解准确率提升，验证了转译模型的有效性。",
                         "bullets": [{"text": "测试人数 42 人"},
                                     {"text": "对照组为原始动线"},
                                     {"text": "误差范围 ±4%"}]}},
            {"id": "s13", "layout": "image-full",
             "content": {"title": "从文本到空间", "subtitle": "一处转折节点的设计推演",
                         "tag": "案例", "image": "", "overlay": 0.55}},
            {"id": "s14", "layout": "closing",
             "content": {"title": "谢谢观看", "subtitle": "恳请各位老师批评指正",
                         "contact": "张三 · 工业设计 · zhang@example.com"}},
        ],
    })
