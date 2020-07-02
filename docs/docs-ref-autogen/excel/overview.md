---
title: Office 脚本 API 参考
description: Office 脚本 JavaScript Api 概述。
ms.date: 06/29/2020
ms.openlocfilehash: 7c4fe97ca35cfb442ebbf9db2e0b03b389185ae8
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/01/2020
ms.locfileid: "45004729"
---
# <a name="office-scripts-api-reference"></a><span data-ttu-id="fc0c4-103">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="fc0c4-103">Office Scripts API reference</span></span>

<span data-ttu-id="fc0c4-104">Office 脚本 API 允许您在 web 上的 Excel 中自动执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="fc0c4-105">使用此参考文档可了解有关可用于脚本的类、方法和其他类型的详细信息。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="fc0c4-106">可通过 Office 脚本访问的所有对象都可以在页面左侧的目录中找到。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="fc0c4-107">如果你要查找用于开发 Office 外接程序的 JavaScript Api，请访问[Office 外接程序 JAVASCRIPT api 参考](/javascript/api/overview?view=excel-js-preview)。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="fc0c4-108">公共类</span><span class="sxs-lookup"><span data-stu-id="fc0c4-108">Common classes</span></span>

<span data-ttu-id="fc0c4-109">以下列表细分了 Office 脚本对象模型的基础知识。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="fc0c4-110">这将显示常见类以及它们之间的关系。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="fc0c4-111">一个 [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) 包含一个或多个 [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-111">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excelscript/excelscript.worksheet).</span></span>
- <span data-ttu-id="fc0c4-112">[Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) 可通过 [Range](/javascript/api/office-scripts/excelscript/excelscript.range) 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-112">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excelscript/excelscript.range) objects.</span></span>
- <span data-ttu-id="fc0c4-113">[Range](/javascript/api/office-scripts/excelscript/excelscript.range) 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-113">A [Range](/javascript/api/office-scripts/excelscript/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="fc0c4-114">[Range](/javascript/api/office-scripts/excelscript/excelscript.range) 用于创建和放置 [Table](/javascript/api/office-scripts/excelscript/excelscript.table)、[Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) 和 [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape) 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-114">[Ranges](/javascript/api/office-scripts/excelscript/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excelscript/excelscript.table), [Charts](/javascript/api/office-scripts/excelscript/excelscript.chart), [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="fc0c4-115">[工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet)包含使用单个工作表中存在的那些对象填充的数组。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-115">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="fc0c4-116">[工作簿](/javascript/api/office-scripts/excelscript/excelscript.workbook)包含所有工作簿的一些数据对象的数组。</span><span class="sxs-lookup"><span data-stu-id="fc0c4-116">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="fc0c4-117">有关 Office 脚本对象模型的详细信息，请访问[web 上 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="fc0c4-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="fc0c4-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fc0c4-118">See also</span></span>

- [<span data-ttu-id="fc0c4-119">关于 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="fc0c4-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="fc0c4-120">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="fc0c4-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="fc0c4-121">Excel 网页版中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="fc0c4-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
