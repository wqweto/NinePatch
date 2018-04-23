## NPPNG
Nine Patch PNGs for VB6

### Description

NPPNG project includes `cNinePatch.cls` for 9-patch parsing/handling and `ctxNineButton.ctl` custom user-control for transparent animated buttons with states based on 9-patches.

### Supported UI through `Style` property

| Type            | Style                      | Preview                                   |
| --------------- | -------------------------- | ----------------------------------------- |
| Buttons         | Simple buttons                                                        ||
|                 | `ucsBtyButtonDefault`      | ![ButtonDefault](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/ButtonDefault.png)   |
|                 | `ucsBtyButtonGreen`        | ![ButtonGreen](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/ButtonGreen.png)       |
|                 | `ucsBtyButtonTurnRed`      | ![ButtonTurnRed](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/ButtonTurnRed.png)   |
| Flat Buttons    | Flat buttons based on bootstrap                                       ||
|                 | `ucsBtyFlatPrimary`        | ![FlatPrimary](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatPrimary.png)       |
|                 | `ucsBtyFlatSecondary`      | ![FlatSecondary](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatSecondary.png)   |
|                 | `ucsBtyFlatSuccess`        | ![FlatSuccess](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatSuccess.png)       |
|                 | `ucsBtyFlatDanger`         | ![FlatDanger](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatDanger.png)         |
|                 | `ucsBtyFlatWarning`        | ![FlatWarning](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatWarning.png)       |
|                 | `ucsBtyFlatInfo`           | ![FlatInfo](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatInfo.png)             |
|                 | `ucsBtyFlatLight`          | ![FlatLight](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatLight.png)           |
|                 | `ucsBtyFlatDark`           | ![FlatDark](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/FlatDark.png)             |
| Outline Buttons | Outline buttons based on boostrap                                     ||
|                 | `ucsBtyOutlinePrimary`     | ![OutlinePrimary](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlinePrimary.png) |
|                 | `ucsBtyOutlineSecondary`   | ![OutlineSecondary](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineSecondary.png) |
|                 | `ucsBtyOutlineSuccess`     | ![OutlineSuccess](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineSuccess.png) |
|                 | `ucsBtyOutlineDanger`      | ![OutlineDanger](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineDanger.png)   |
|                 | `ucsBtyOutlineWarning`     | ![OutlineWarning](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineWarning.png) |
|                 | `ucsBtyOutlineInfo`        | ![OutlineInfo](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineInfo.png)       |
|                 | `ucsBtyOutlineLight`       | ![OutlineLight](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineLight.png)     |
|                 | `ucsBtyOutlineDark`        | ![OutlineDark](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/OutlineDark.png)       |
| Cards           | Static cards that can be owner-drawn with information                 ||
|                 | `ucsBtyCardDefault`        | ![CardDefault](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/CardDefault.png)       |
|                 | `ucsBtyCardPrimary`        | ![CardPrimary](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/CardPrimary.png)       |
|                 | `ucsBtyCardSuccess`        | ![CardSuccess](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/CardSuccess.png)       |
|                 | `ucsBtyCardPurple`         | ![CardPurple](https://raw.githubusercontent.com/wqweto/NinePatch/master/doc/CardPurple.png)         |

### Supported states

| State                | Description |
| -------------------- | ----------- |
| `ucsBstNormal`       | Normal state |
| `ucsBstHover`        | Optional mouse hover. Otherwise `ucsBstNormal` is used |
| `ucsBstPressed`      | Optional mouse pressed. Otherwise `ucsBstHover` is used |
| `ucsBstHoverPressed` | Optional mouse pressed over the control. Otherwise `ucsBstPressed` is used |
| `ucsBstDisabled`     | Optional disabled state |
| `ucsBstFocused`      | Optional focus rectangle which is composited before current state |

Each state can have a different 9-patch image, text and shadow properties.

### Supported state customization

| Property              | Description |
| --------------------- | ----------- |
| `ButtonImageArray`    | Nine-patch image as byte array |
| `ButtonImageOpacity`  | Image transparency incl. composited text and shadow |
| `ButtonTextFont`      | Text font |
| `ButtonTextFlags`     | Text alignment and wrapping |
| `ButtonTextColor`     | Text color |
| `ButtonTextOpacity`   | Text transparency on image before applying image transparency |
| `ButtonTextOffsetX`   | Additional text horizontal offset after alignment |
| `ButtonTextOffsetY`   | Additional text vertical offset after alignment |
| `ButtonShadowColor`   | Text shadow color |
| `ButtonShadowOpacity` | Text shadow transparency on image |
| `ButtonShadowOffsetX` | Shadow horizontal offset from text |
| `ButtonShadowOffsetY` | Shadow vertical offset from text |

### Global customizations

| Property            | Description |
| ------------------- | ----------- |
| `AnimationDuration` | Duration of state transition (in seconds) |
| `Opacity`           | Global control transparency |
| `Enabled`           | Uses disabled state settings to paint control |
