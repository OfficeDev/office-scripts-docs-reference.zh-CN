---
title: Office 脚本 API 参考
description: Office 脚本异步 JavaScript Api 概述。
ms.date: 06/29/2020
ms.openlocfilehash: 3d8d37b30d9535e8b6a56a08c44f9034cb599f31
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/01/2020
ms.locfileid: "45003952"
---
# <a name="office-scripts-async-api-reference"></a>Office 脚本异步 API 参考

Office 脚本异步 API 支持在 Office 脚本预览阶段执行较旧的脚本。 使用此参考文档可详细了解这些旧脚本使用的类、方法和其他类型。 可通过 Office 脚本访问的所有对象都可以在页面左侧的目录中找到。

> [!IMPORTANT]
> 强烈建议使用标准 Office 脚本 Api 创建新的脚本。 如果要进行或编辑新脚本，请切换到[非异步版本](?view=office-scripts)的 api。

## <a name="common-classes"></a>公共类

以下列表细分了 Office 脚本对象模型的基础知识。 这将显示常见类以及它们之间的关系。

- [工作簿](/javascript/api/office-scripts/excelscript/excelscript.workbook)在[WorksheetCollection](/javascript/api/office-scripts/excelscript/excelscript.worksheetcollection)中包含一个或多个[工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet)。
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) 可通过 [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 对象访问单元格。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 代表一组连续的单元格。
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 用于创建和放置 [Table](/javascript/api/office-scripts/excelscript/excelscript.table)、[Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) 和 [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape) 以及其他数据可视化或组织对象。
- [工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet)包含单个工作表中存在的这些数据对象（如[ChartCollection](/javascript/api/office-scripts/excelscript/excelscript.chartcollection)）的集合。
- [工作簿](/javascript/api/office-scripts/excelscript/excelscript.workbook)包含整个[工作簿](/javascript/api/office-scripts/excelscript/excelscript.workbook)的一些数据对象（如[TableCollection](/javascript/api/office-scripts/excelscript/excelscript.tablecollection)）的集合。

有关 Office 脚本对象模型的详细信息，请访问[web 上 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>另请参阅

- [使用 Office 脚本异步 Api 支持旧版脚本](/office/dev/scripts/develop/excel-async-model)
- [关于 Office 脚本](/office/dev/scripts/overview/excel)
- [在 Excel 网页版中录制、编辑和创建 Office 脚本](/office/dev/scripts/tutorials/excel-tutorial)
- [Excel 网页版中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)
