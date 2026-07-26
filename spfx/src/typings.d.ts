declare module "*.module.scss" {
  const styles: { [className: string]: string };
  export default styles;
}

declare module "*.png" {
  const url: string;
  export default url;
}
