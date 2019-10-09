// @todo test this
export function timeout<T>(promise: PromiseLike<T>, timeoutMs: number, message?: string) {
    return new Promise<T>((resolve, reject) => {
        const timeoutHandle = setTimeout(() => {
            reject(message == null ? `Timed out after ${timeoutMs} ms.` : message);
        }, timeoutMs);

        // Maybe use finally when it's available.
        promise.then(
            result => {
                resolve(result);
                clearTimeout(timeoutHandle);
            },
            reason => {
                reject(reason);
                clearTimeout(timeoutHandle);
            },
        );
    });
}

export interface PromiseResult<T = any> {
    state: "fulfilled" | "rejected";
    value?: T;
    reason?: any;
}

export function allSettled<T = any>(promises: PromiseLike<T>[]): Promise<PromiseResult<T>[]> {
    const results = new Array(promises.length);
    return new Promise(resolve => {
        let count = 0;
        for (let i = 0 ; i < promises.length; ++i) {
            const promise = promises[i];
            promise.then((result) => {
                results[i] = {
                    state: "fulfilled",
                    value: result
                }
            }, (reason) => {
                results[i] = {
                    state: "rejected",
                    reason: reason
                }
            }).then(() => {
                if (++count === promises.length) {
                    resolve(results);
                }
            });
        }
    });
}

export function realPromise<T>(promise: PromiseLike<T>): Promise<T> {
    return new Promise((resolve, reject) => {
        promise.then(resolve, reject);
    });
}