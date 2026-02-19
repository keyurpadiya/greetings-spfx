import * as React from 'react';
import styles from './Greetings.module.scss';
import type { IGreetingsProps } from './IGreetingsProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface IGreetingsState {
  imageBrightness: number;
  autoFontColor: string; // Renamed from fontColor to autoFontColor
  isLoading: boolean;
}

export default class Greetings extends React.Component<IGreetingsProps, IGreetingsState> {
  private brightnessCache = new Map<string, number>();

  constructor(props: IGreetingsProps) {
    super(props);

    this.state = {
      imageBrightness: 128,
      autoFontColor: '#000000', // Renamed
      isLoading: false
    };
  }

  private getUserName(): string {
    const { userDisplayName, showFirstNameOnly } = this.props;

    if (!userDisplayName) return 'User';

    if (showFirstNameOnly) {
      const firstName = userDisplayName.split(' ')[0];
      return firstName;
    }

    return userDisplayName;
  }

  private getTimeBasedGreeting(): string {
    const hour = new Date().getHours();

    if (hour < 12) {
      return "Good Morning";
    } else if (hour < 17) {
      return "Good Afternoon";
    } else {
      return "Good Evening";
    }
  }

  private calculateFontSizeDetails(fontSize: string): {
    value: number;
    unit: string;
    isValid: boolean
  } {
    if (!fontSize) {
      return { value: 12, unit: 'px', isValid: false };
    }

    const fontSizeRegex = /^(\d+(\.\d+)?)(px|rem|em|pt|%)$/;
    const match = fontSize.match(fontSizeRegex);

    if (match) {
      return {
        value: parseFloat(match[1]),
        unit: match[3],
        isValid: true
      };
    }

    return { value: 12, unit: 'px', isValid: false };
  }

  private getSizeClass(fontSize: string): string {
    const fontSizeDetails = this.calculateFontSizeDetails(fontSize);

    if (!fontSizeDetails.isValid) {
      return styles.mediumFont;
    }

    const { value, unit } = fontSizeDetails;

    let pixelValue = value;
    if (unit === 'rem' || unit === 'em') {
      pixelValue = value * 16;
    } else if (unit === 'pt') {
      pixelValue = value * 1.333;
    }

    if (pixelValue < 14) {
      return styles.smallFont;
    } else if (pixelValue >= 14 && pixelValue < 18) {
      return styles.mediumFont;
    } else if (pixelValue >= 18 && pixelValue < 24) {
      return styles.largeFont;
    } else {
      return styles.extraLargeFont;
    }
  }

  private async getImageBrightness(imageUrl: string): Promise<number> {
    if (this.brightnessCache.has(imageUrl)) {
      return this.brightnessCache.get(imageUrl)!;
    }

    if (!imageUrl || imageUrl.trim() === '') {
      return 128;
    }

    return new Promise<number>((resolve) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';

      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          const maxSize = 100;
          const ratio = Math.min(maxSize / img.width, maxSize / img.height);

          canvas.width = Math.floor(img.width * ratio);
          canvas.height = Math.floor(img.height * ratio);

          const ctx = canvas.getContext('2d');
          if (!ctx) {
            const defaultBrightness = 128;
            this.brightnessCache.set(imageUrl, defaultBrightness);
            resolve(defaultBrightness);
            return;
          }

          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

          const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
          const data = imageData.data;

          let totalBrightness = 0;
          let pixelCount = 0;

          for (let i = 0; i < data.length; i += 16) {
            const brightness = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
            totalBrightness += brightness;
            pixelCount++;
          }

          const averageBrightness = pixelCount > 0 ? totalBrightness / pixelCount : 128;
          this.brightnessCache.set(imageUrl, averageBrightness);
          resolve(averageBrightness);
        } catch (error) {
          console.log('Error analyzing image brightness:', error);
          const defaultBrightness = 128;
          this.brightnessCache.set(imageUrl, defaultBrightness);
          resolve(defaultBrightness);
        }
      };

      img.onerror = () => {
        console.log('Error loading image for brightness detection');
        const defaultBrightness = 128;
        this.brightnessCache.set(imageUrl, defaultBrightness);
        resolve(defaultBrightness);
      };

      img.src = imageUrl;

      setTimeout(() => {
        if (!img.complete) {
          const defaultBrightness = 128;
          this.brightnessCache.set(imageUrl, defaultBrightness);
          resolve(defaultBrightness);
        }
      }, 2000);
    });
  }


  private getTextShadow(brightness: number): string {
    if (brightness < 128) {
      return '1px 1px 3px rgba(0, 0, 0, 0.7), 0 0 5px rgba(0, 0, 0, 0.5)';
    } else {
      return '1px 1px 2px rgba(255, 255, 255, 0.7)';
    }
  }

  private getFontStyle(): React.CSSProperties {
    const { fontSize, fontStyle, fontColor } = this.props;
    const { autoFontColor, imageBrightness } = this.state;

    const styleObj: React.CSSProperties = {
      fontWeight: 'normal',
      fontStyle: 'normal',
      // ✅ Priority: Manual color > Auto color > Default
      color: fontColor && fontColor.trim() !== '' ? fontColor : (autoFontColor || '#000000'),
      textShadow: this.getTextShadow(imageBrightness)
    };

    const fontSizeDetails = this.calculateFontSizeDetails(fontSize);
    styleObj.fontSize = fontSizeDetails.isValid ? fontSize : '12px';

    switch (fontStyle) {
      case 'Italic':
        styleObj.fontStyle = 'italic';
        styleObj.fontWeight = 'normal';
        break;
      case 'Bold':
        styleObj.fontWeight = 'bold';
        styleObj.fontStyle = 'normal';
        break;
      case 'Bold Italic':
        styleObj.fontWeight = 'bold';
        styleObj.fontStyle = 'italic';
        break;
      default:
        break;
    }

    return styleObj;
  }
  private getContainerWidth(): number {
    const element = this.props.context?.domElement;
    if (element && element.parentElement) {
      return element.parentElement.clientWidth;
    }
    return window.innerWidth;
  }


  private getWelcomeContentStyle(): React.CSSProperties {
    const { textAlignment } = this.props;

    const styleObj: React.CSSProperties = {};

    // Adjust flex based on alignment
    switch (textAlignment) {
      case 'left':
        styleObj.justifyContent = 'flex-start';
        break;
      case 'right':
        styleObj.justifyContent = 'flex-end';
        break;
      case 'center':
      default:
        styleObj.justifyContent = 'center';
        break;
    }

    return styleObj;
  }

  private getTextAlignment(): React.CSSProperties {
    const { textAlignment } = this.props;

    const styleObj: React.CSSProperties = {
      textAlign: (textAlignment as any) || 'center' // Default to center
    };

    return styleObj;
  }

  private getWelcomeBoxStyle(): React.CSSProperties {
    const { backgroundImageUrl } = this.props;
    const containerWidth = this.getContainerWidth();
    const isFullWidthLayout = containerWidth > 900;

    const styleObj: React.CSSProperties = {
      position: 'relative',
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      width: '100%',
      maxWidth: isFullWidthLayout ? '100%' : '420px',
      minHeight: '100px', // ✅ INCREASED from 70px to 100px (client webpart ki height)
      padding: '20px 28px', // ✅ INCREASED padding for better spacing
      margin: '0',
      transition: 'all 0.3s ease',
      backgroundSize: 'cover',
      backgroundPosition: 'center center', // ✅ SPECIFIC center positioning
      backgroundRepeat: 'no-repeat',
      backgroundAttachment: 'local', // ✅ Added for better image display
      flexShrink: 0

    };

    let urlString = '';

    if (!urlString) {
      styleObj.backgroundColor = 'transparent';
    }

    if (backgroundImageUrl) {
      if (typeof backgroundImageUrl === 'string') {
        urlString = backgroundImageUrl;
      } else if (backgroundImageUrl.fileAbsoluteUrl) {
        urlString = backgroundImageUrl.fileAbsoluteUrl;
      } else if (backgroundImageUrl.spItemUrl) {
        urlString = backgroundImageUrl.spItemUrl;
      }
    }

    if (urlString?.trim()) {
      styleObj.backgroundImage = `url('${encodeURI(urlString.trim())}')`;
      // ✅ Use 'contain' for some images to prevent cropping
      // You can make this configurable or detect based on image aspect ratio
      const shouldUseContain = this.shouldUseContainBackground(urlString);
      styleObj.backgroundSize = shouldUseContain ? 'contain' : 'cover';
    }

    return styleObj;
  }
  private shouldUseContainBackground(imageUrl: string): boolean {
    // You can add logic here based on image dimensions or aspect ratio
    // For now, we'll use 'cover' for all images as client webpart does
    return false;

    // Future enhancement: Check image dimensions
    // return new Promise<boolean>((resolve) => {
    //   const img = new Image();
    //   img.onload = () => {
    //     const aspectRatio = img.width / img.height;
    //     // If image is very wide or very tall, use 'contain'
    //     resolve(aspectRatio > 2 || aspectRatio < 0.5);
    //   };
    //   img.src = imageUrl;
    // });
  }

  private extractImageUrl(): string {
    const { backgroundImageUrl } = this.props;

    if (!backgroundImageUrl) return '';

    if (typeof backgroundImageUrl === 'string') {
      return backgroundImageUrl.trim();
    } else if (backgroundImageUrl.fileAbsoluteUrl) {
      return backgroundImageUrl.fileAbsoluteUrl;
    } else if (backgroundImageUrl.spItemUrl) {
      return backgroundImageUrl.spItemUrl;
    }

    return '';
  }

  private hasBackgroundImage(): boolean {
    const { backgroundImageUrl } = this.props;

    if (!backgroundImageUrl) return false;

    if (typeof backgroundImageUrl === 'string') {
      return backgroundImageUrl.trim() !== '';
    }

    return !!(backgroundImageUrl.fileAbsoluteUrl || backgroundImageUrl.spItemUrl);
  }

  private determineAutoFontColor(brightness: number): string { // Renamed
    const DARK_THRESHOLD = 128;
    const VERY_DARK_THRESHOLD = 64;
    const VERY_LIGHT_THRESHOLD = 192;

    if (brightness < VERY_DARK_THRESHOLD) {
      return '#FFFFFF';
    } else if (brightness < DARK_THRESHOLD) {
      return '#F0F0F0';
    } else if (brightness > VERY_LIGHT_THRESHOLD) {
      return '#000000';
    } else {
      return '#333333';
    }
  }

  private async updateImageBrightness(): Promise<void> {
    const imageUrl = this.extractImageUrl();

    if (!imageUrl) {
      this.setState({
        imageBrightness: 255,
        autoFontColor: '#000000', // Updated
        isLoading: false
      });
      return;
    }

    this.setState({ isLoading: true });

    try {
      const brightness = await this.getImageBrightness(imageUrl);
      const autoFontColor = this.determineAutoFontColor(brightness);
      this.setState({
        imageBrightness: brightness,
        autoFontColor: autoFontColor, // Updated
        isLoading: false
      });
    } catch (error) {
      console.log('Error detecting image brightness:', error);
      this.setState({
        imageBrightness: 128,
        autoFontColor: '#000000', // Updated
        isLoading: false
      });
    }
  }

  componentDidMount(): void {
    this.updateImageBrightness();
  }

  componentDidUpdate(prevProps: IGreetingsProps): void {
    // Update brightness if background image changed
    if (prevProps.backgroundImageUrl !== this.props.backgroundImageUrl) {
      this.updateImageBrightness();
    }

    // ✅ NEW: Clear brightness cache if manual color is set/cleared
    if (prevProps.fontColor !== this.props.fontColor) {
      // If user sets manual color, we might want to clear cache
      // or if they clear it, we need to recalculate
      if ((!this.props.fontColor || this.props.fontColor.trim() === '') && this.props.backgroundImageUrl) {
        this.updateImageBrightness();
      }
    }
  }
  public render(): React.ReactElement<IGreetingsProps> {
    const {
      hasTeamsContext,
      greetingText,
      fontSize
    } = this.props;

    const userName = this.getUserName();
    const fontStyle = this.getFontStyle();
    const welcomeBoxStyle = this.getWelcomeBoxStyle();
    const sizeClass = this.getSizeClass(fontSize);
    const hasBackground = this.hasBackgroundImage();
    const textAlignmentStyle = this.getTextAlignment();
    const welcomeContentStyle = this.getWelcomeContentStyle();


    const combinedFontStyle = {
      ...fontStyle,
      ...textAlignmentStyle
    };

    const welcomeBoxClass = `${styles.welcome} ${sizeClass} ${hasBackground ? '' : styles.noBackgroundImage
      }`;

    return (
      <section
        className={`${styles.greetings} ${hasTeamsContext ? styles.teams : ''}`}
      >
        <div className={styles.contentContainer}>
          <div
            className={welcomeBoxClass}
            style={welcomeBoxStyle}
          >


            <div className={styles.welcomeContent} style={welcomeContentStyle}>
              <h2 style={combinedFontStyle}>
                {escape(greetingText || this.getTimeBasedGreeting())}, {escape(userName)}!
              </h2>
            </div>
          </div>
        </div>
      </section>
    );
  }
}