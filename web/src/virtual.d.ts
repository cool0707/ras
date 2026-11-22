declare module "virtual:samples" {
    const samples: { id: string; title: string; files: { path: string; text: string }[] }[];
    export default samples;
}
