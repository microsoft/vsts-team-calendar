import React = require("react");

import { Dialog } from "azure-devops-ui/Dialog";

interface IMessageDialogProps {
    message: string;

    /**
     * Callback function on dialog dismiss
     */
    onConfirm: () => void;

    /**
     * Callback function on dialog dismiss
     */
    onDismiss: () => void;

    title: string;
}

/**
 * Dialog that lets user view scope changes for extension
 * only admin can authorize extension update
 */
export class MessageDialog extends React.Component<IMessageDialogProps> {
    constructor(props: IMessageDialogProps) {
        super(props);
    }

    public render(): JSX.Element {
        return (
            <Dialog
                titleProps={{ text: this.props.title }}
                footerButtonProps={[
                    {
                        text: "Cancel",
                        onClick: this.props.onDismiss
                    },
                    {
                        text: "Delete",
                        onClick: this.props.onConfirm,
                        primary: true
                    }
                ]}
                onDismiss={this.props.onDismiss}
            >
                {this.props.message}
            </Dialog>
        );
    }
}
