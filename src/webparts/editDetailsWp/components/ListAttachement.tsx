import * as React from 'react';
import { ListItemAttachments } from "@pnp/spfx-controls-react/lib";

export interface ISpfxAttachmentControlProps {
    SeletedList: string;
    SelectedItem: number;
    context: any | null;
}

export class SpfxAttachmentControl extends React.Component<ISpfxAttachmentControlProps, {}> {
    public render(): React.ReactElement<ISpfxAttachmentControlProps> {
        return (
            <div>
                
                    <div><label>Attachments</label>
                        <ListItemAttachments listId={this.props.SeletedList}
                            itemId={this.props.SelectedItem}
                            context={this.props.context}
                            disabled={false} /></div>
                
            </div>
        );
    }
}