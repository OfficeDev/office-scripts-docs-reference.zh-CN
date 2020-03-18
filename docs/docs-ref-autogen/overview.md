---
title: Office 脚本 API 参考
description: Office 脚本 JavaScript Api 概述
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42702539"
---
# <a name="office-scripts-api-reference"></a>Office 脚本 API 参考

Office 脚本 API 允许您在 web 上的 Excel 中自动执行常见任务。 使用此参考文档可了解有关可用于脚本的类、方法和其他类型的详细信息。 可通过 Office 脚本访问的所有对象都可以在页面左侧的目录中找到。

## <a name="common-classes"></a>公共类

以下列表细分了 Office 脚本对象模型的基础知识。 这将显示常见类以及它们之间的关系。

- [工作簿](/javascript/api/office-scripts/excel/excel.workbook)在[WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection)中包含一个或多个[工作表](/javascript/api/office-scripts/excel/excel.worksheet)。
- [工作表](/javascript/api/office-scripts/excel/excel.worksheet)通过[Range](/javascript/api/office-scripts/excel/excel.range)对象提供对单元格的访问。
- [区域](/javascript/api/office-scripts/excel/excel.range)代表一组连续的单元格。
- [区域](/javascript/api/office-scripts/excel/excel.range)用于创建和放置[表](/javascript/api/office-scripts/excel/excel.table)、[图表](/javascript/api/office-scripts/excel/excel.chart)、[形状](/javascript/api/office-scripts/excel/excel.shape)以及其他数据可视化或组织对象。
- [工作表](/javascript/api/office-scripts/excel/excel.worksheet)包含单个工作表中存在的这些数据对象（如[ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)）的集合。
- [工作簿](/javascript/api/office-scripts/excel/excel.workbook)。 包含整个[工作簿](/javascript/api/office-scripts/excel/excel.workbook)的一些数据对象（如[TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)）的集合。

有关 Office 脚本对象模型的详细信息，请访问[web 上 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>另请参阅

- [关于 Office 脚本](/office/dev/scripts/overview/excel)
- [在 Excel 网页上记录、编辑和创建 Office 脚本](/office/dev/scripts/tutorials/excel-tutorial)
- [Web 上的 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)
