import * as React from 'react';
import { useMemo } from 'react';
import * as Color from 'color';

export default function useColorStyle(colorString: string, cssPropName: 'color' | 'backgroundColor'): React.CSSProperties {
  return useMemo<React.CSSProperties>(() => {
    let style: React.CSSProperties = null;
    try {
      if (colorString) {
        const textColor = new Color(colorString);
        if (textColor) {
          style = {};
          style[cssPropName] = colorString;
        }
      }
    }
    finally {
      return style;
    }
  }, [colorString]);
}
