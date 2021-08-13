---
title: Office 脚本 API 参考
description: Office脚本 JavaScript API 的概述。
ms.date: 06/29/2020
ms.openlocfilehash: 3ce3344fb49b2811719feb13f8fb4118f1a20060db9d85a06d1be939f22bf3c5
ms.sourcegitcommit: 6a0182075a558c4fe664fedfee08fea76513b192
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2021
ms.locfileid: "58183424"
---
# <a name="office-scripts-api-reference"></a>Office 脚本 API 参考

借助 Office 脚本 API，可以自动执行 Excel web 版 中的常见Excel web 版。 使用此参考文档可了解有关可用于脚本的类、方法和其他类型的信息。 所有可通过 Office Scripts 访问的对象都可以在页面左侧的目录下找到。

> [!NOTE]
> 若要查找用于开发加载项Office JavaScript API，请访问 Office[加载项 JavaScript API 参考](/javascript/api/overview?view=excel-js-preview)。

## <a name="common-classes"></a>常用类

下面的列表对脚本对象Office的基础知识。 这将显示常用类及其相互关联。

- 一个 [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) 包含一个或多个 [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)。
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) 可通过 [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 对象访问单元格。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 代表一组连续的单元格。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 用于创建和放置 [Table](/javascript/api/office-scripts/excelscript/excelscript.table)、[Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) 和 [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape) 以及其他数据可视化或组织对象。
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)包含用单个工作表中的对象填充的数组。
- 工作簿 [包含](/javascript/api/office-scripts/excelscript/excelscript.workbook) 整个 Workbook 的其中一些数据对象的数组。

有关脚本脚本对象Office，请参阅脚本[的脚本基础Office脚本Excel web 版](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>另请参阅

- [关于Office脚本](/office/dev/scripts/overview/excel)
- [在 Excel 网页版中录制、编辑和创建 Office 脚本](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel 网页版中 Office 脚本的脚本基础知识](/office/dev/scripts/develop/scripting-fundamentals)
