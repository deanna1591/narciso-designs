import React, { useState, useEffect, useMemo, useRef } from "react";
import { supabase } from "./lib/supabaseClient";
import { store } from "./lib/store";
import * as XLSX from "xlsx";
import {
  Plus, RefreshCw, Search, Link2, Check, X, ChevronDown, ChevronRight,
  Eye, EyeOff, Trash2, AlertTriangle, ExternalLink, Package, Loader2,
  ArrowUp, ArrowDown, Image as ImageIcon, RotateCcw, List, LayoutGrid, Percent,
  ChevronLeft, Truck, Paperclip, FileText, Users, Building2, MapPin, Inbox, Home, Folder, Upload, Pin,
  Presentation, Download, CheckSquare, Square, LogOut, Lock, Copy, KeyRound, Bell, MessageSquare, Wallet, CreditCard
} from "lucide-react";

const LOGO = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAlgCWAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACmAKcDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKoa5rdp4d0u41C9lEVvCpZmP8h70AX6K8W0n4zX91YXmtNb+ZaLPsW1A+bZnGR717FY3aX1pDcR5CSKGAPUZHQ1KkmBPRRRVAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAleJftGa9DdeH7KztbuGZDPukWKQE5HQHH1Nbnxm+JVx4LmsLGCHzEu0ZpWzhtvQgehrxm88VeGmhmeHSpGuHUj942QCe4rKclsUkangfVE0bwNdXjx+cscpOz1r3H4Za1bap4aS8S5BSR/us/3OB8vtXiFn4a1Oz+D9xfGzYwTylgo+8I/wC+R6VzHgvxlJ4XaaGUPNYzD5o1bGD6io5uWw7XPsgEMAQcilryb4G+Nm15b/TQ0r21th7czHLKp6rn0GeK9ZrdO6uQFFFFMAooooAKKKKACiiigAooooAKKKKACisnVPFGn6NqWn2F1OsVxfMVhDcAkD1rWoAKazCNSzEKoGST0FcF8Sfi7p/gHZbLH9u1J8EW6tgKvqx7VoeNtTlvfhvqF7ZZzLaFuOoUj5v0zU8yA8T+OfjzSPF9/b2umxvLJZuVa8zhW65UDuM964zwF4bPivxbp2nMCYXkDS/7o5NdH4H+COteMLGO+kdbCyk5Rps73H94D/GvVfhr8IZfh3rd5qVxeR3UXkFFwDkdCT+hrDlcndmmiWhu/wDCaabL4yfwX5KeT9lKn0zj7uPp/Ovmjx14dbwr4qv9NxhI3zH/ALp5H6Vei8TTL8Tf7ZZ9z/bsk/3hu2gfyr2b4qfCUeMtcttVGo2+m2/k7JmmOCW7EHp0pv31oJaM4X4F+PtI8J3VxZ38bQy3jqBdZyo9iOw6V9KKwkUMp3KwyCO9fKfjT4L6x4T086hBJHqViOfNgOSo9T7fSvoHwrq7ab8NNP1G8DO0Nisr88nA/nVwb2YpeR1tFcH8Ofi1p/j7zLfy/sWoR5Jt2bO4Z6g967utE77Ei0VkaT4p07W9Qv7G0nWW5smCTKO1a9MAooooAKKKKACiiigAoqK5uobOFpp5UhiX7zyMAB+NU9M8Q6brLMtjfQ3LL1WNwSPwoA8Z+N980fjGxs9RVl0u4iTyLheGgmDH5lP5Zr0K48eaf4QtYNM1a/L6nHaeZ5rIQsxA6g9Mn0rw/wCPFpqGm+Nm+1XM09hKVmt1kYssY7qPTp+tbHxM1S11mbSPt5xp2pWUcltc7eYHAwc+xNYc1myjybV/EA8Razf3U95FLqDSb3t9/wA6g9Dj0r3bwV8aLS+m0vw4ukzXVo0QgkkxuOe5x/dr5V+O3wz13wff+GvGnhyM3WqrE8N7ZRci7hDZBX1OO1eyfDf45eF7XwJaR+D3hm8UTRtLqUV4uy4tGHVWVueO3akk1qPfQ7/43eN9a0+9XSdFSeysYEHmzQIRuJHQY7AVw3hD4xazofmW97dSX+nzRsjLIcsmRgFTXb/Cf4van4m1xdF14RXsVypCSGMA5/ukDjFcT+1la+Gvg/oY8SRTRw3V1J5UOjof3lzIenlr1x69hRrL3kGi0Z478QfiNYeALT7fOPtF5LN/otmh+ed85AA9PU1jHRPEvxauP7f+I2r3tzczYNvo0Fw0dtaR/wAKbAeT9a5rwV4Xnn1F/HPjyaKDUGGbSzuGCpZx9uD/ABV7F4K0zX/itIbHwXp8hSf5JvE19GVs7KPozRA8zSemOBSV9kGm50/7EusXlnqXjLwxLJeaj4U+3GGw+1M0yxOEBkQFv4favVfGnxrtov7Z8Oy6TLa27QSW8cjDackEZ2+nNchqfiI/BKxtPAvhAC0tdLX9/dzANNczN8zyOSOSxOa6JbkfGPwzeWeqWcdv4ks4Bc208YwZ0IOPwOKty6Jit1PJfD97deHbqz1u2ZgIJwpK9eOx+or6ek+IemeJ9Nv9P0m936obBpl8sHCkr0B9RXzVpEAk8Ia+snWGWFh7Nkiuk+Da/YZte1dwVhs7ByGPQuQQF+vNZxdtCmupufs4297J4u1KYszxRxlbhzn5mJOM++c19H18wfAa31HUvHHmW1xNDYREz3CIxCOTnAYd+9fR2p+ItM0UoL6+htmY4CyOAfyrWnsQzRoqK3uIrqISQyLLG3RkIIP41LWogooooAKKKKAPA/jBfav448aQeE9IZmijUNIqthS3ct9K4GHT774aeO9Pit9Tt7q5EypJ9jckDJwUYetes/DqETfGTxs84/eIQI/YE4OKv6f8CtM0/wAXS67cXjzxLL9oSFwAFbryfasOVvVFXsbnxO+H8fxA8PiNQIr+Eb4JG9ccqfY14vDZNd+Gbrwf4ldNJ1XT2a40+e7YIjjum48c9q7nSvEuufGpr3UNN1Z/C/gK3mkt4r23A+16kUYq7q54iiyCARycZ6Vx3xP8J/C/Q7PZPpSeMNbkHyyapdSXRX3clvTpVyS3YlfYwPC/ibR9f0G20nUNRt5bnSbvdFLHMrtGjEeh5XI5+tWPil8GfA3jj4vSReJbL+z5L6yRbbVtPkNvJHJjhty4DH65rzl/hz8MPElr9n1PwrFoV2v3dR0NnhOPR1U8/hVu6+BPxJj8Kyap4L8XnxLo9gGe203W5N5KrklIpD8yY9DwaheQxfFX7P8A4k+AqS6vo3xgureSAF7aG+sVuX2gZyzE4Ucda8G8E/CX4hftIatL4617xFqNzb2czG31BgUyqH76H7qD0x1rUtPi94n/AGjLnSvhvpEt1az3rsNSiklDbYk5ZEbqc4r6g0PUfFvgvw2mgf2alzoFrGIHs0RcLGvG3K8jpTbaA5fwz8M/APhvwzB4nvfDN94n1gSmAXmvXTXMBk6+ZtPy8fStnVNT8ZeHNTtdVmkuLZDteD7PkW+3qFCr8uMdq9sm8UeFB8I/tn9np/ZLJ5YsMDd5nTb9c968Q8deNviA3gC/bSiuiWdjZyT2oeFW+WNCwB38twKT9QO7XwZpXxqvI9cttQNheMqjUbZ15BAxuXI5zitDStQsI/ii8OllW0rR9Le2knDZUYXufzrwf4OeLvHHxMurjwnqTSaL4qtbeO4u9NlcQieBwCs8f95CGHTpmu8j0LVvAvirxX4en1aEaHHY2U9zM0ADBpjLuTcPmK/uxx7mlr2AzDbzzaH9itYGe+1m78xIU5PlqSAfoSTXXapoN3pfh638FaFA17qt0yy6lNGPlU9kLeg9ax1+JVr4ZkL6BaLPfsAp1K9jGQo/hjT+Ff1r1/4YfFfR/Eln5NytvpWpry8fCrJ6sDRFJjd7Gv8ADnwHH8O/DLx8TahIpkuJF/iYDgD2FfPP9k33xK8bags+o29rdSSnZ9scgdeEWvriGaO6hSWJ1kicZVlOQQe9eZXnwI02fxpFrkF3JbxiYXDW6jI3g54PoTVyjskJOxy3wbvdX8E+NLjwhqwYLIm+JS2VGBncD6EV71XkHj+FIfjB4Vmh/wBfIMSY64BGP0r1+qj1QmFFFFWIKKKKAPHvEUx8A/GC01Zx5el6sghlfoofnkmvUNes31TQb+2gfbJPbuiOp7lTiub+La6FJ4PuF16TyYP+WLoMusnYrXHfBj4uRahbw6Dq04W7jxHbXEhx5y9AD/tVndRdh7nlX7Pnx28N+AvBK/C3x1v0fXNHaazYzgLFeRF2IkRiRnIPI9RXR6b8PPA3ifVnfTfHFr9lLbvszMocA/whiee/0qL9pj4N2GoXZ1qbTIb/AE+44nV49xif+9nsD7V86/8ACl/BysJBoyEjkDzpMfluqJPX3h9Lo9h/aA8aeDfh/JpfhvQ5YtSuiu9rTSyLm7mY8DdtzgfXA5rm9Kj+KXxu0uLQIIJvC3g+3XbLpWmybbi6H8X2icdM55Va9c8QeDvC/wAP/g3ZXXhbw7Y6NPfCOOa4t4VEpBHzAv8Ae5I9a6vwj8RtG8Ir4c8O2VmXt7u3WRrqPrvbPUdznrT0TA+a/jN+zLq/hbwjp/inwd4Yg8O+I/CY+1Qz6Sxzcwry6yc5LAZOe/Ipvw3+Jd1rWl2HifTbiSG4uV3yLuON+cMrevINeyeOPj5rGo39zbaUi2WnqTHiRcyP1Bznpn0r5/8Agjeaf4Y8feJPh3rFsp07V2bVtGmX5WikPEkan+ntSfvaIa01Z7xNrEUmj/279nU6cD9pW1I/d/bcbcY9P4sV5F8QNR1HX/D3iC5uppry5NhcOx5OFETE8dgBXodxYS2umN4cuL+KKwhn+2tcE8+XtIxt/vZ7VTn8aQafpU+kaXpFqmnzqUma9hEslwP9rPGD6dKzGWfGHw28b/Fbwv4L1/w/pcPhrxToNlbzaPriyYmZfKXMUmDhon7qRxn65xfhr8Ybz4geMvH2m/EPQn8M6w1pY6ZeLsJjguovOKsPVW3hgfT1613nw5+N2o6FNZaVdWsM2lArDHHboFaIdAFA4wPSug8bfELwnrX/AAmFpqFjbWRht1A1KQBWmdAdgJ4PBPAPqa1Urrcgq/Cz4btq0OtaL4h03NtHte3u8YOWBwyN/EMVU1T4U+E/COsp/a/isLCGLfZAmJCB0G4Hr9a6r4ZeLrmH4NXd8X85rFWjhcnOVwMHP41873FxJeTy3Fw5lmkYs8jHJJ9ah2iloUrs+kb345aLb29vpfh2CW8vHKwQLtwig8bs98eleo2KyW+nwi6kDTLGDLJ0BbHJ+lfP/wCz/wDDx77UV8SXsW20gyLUNxvfu30FdB8Z/i5BZRv4f0m43XEh2XdzHyIl7qPVq0jLS7JLPheT/hPfi9qGtxc6bpKfZ4mI4kfoSPf/AAr2CuN+E6aHH4OtV0KUTQf8tXIw5k/i3fjXZVcQCiiiqEFFFFAHH/FHwKPHvhl7JHEd1E3nQMem4AjB9jmvk/VtH1Dw/fPb31tNZ3MTbfmBGG9j/hX3BVW80uz1Ag3VpDcFeB5sYbH5is5Q5tSk7Hg3wz+JHiHUIv7F1LR5/EFky7TIy/Mq9PmJ4Kjn3rA+K3gXw9o7T3Wg6pbh1b99pvmbmTP92vZ/i3cN4d+HGqPpyLaNtCboVC7QTyeP8814N4O+EOpeKLUalezppGlk5+1XRwWHqAev41nJP4dwR2/hnxx4d8aeGNP8G6ik0MkkYRbgYwsmTiiL4T/8KstbrxLqF8upGxTdawRoVXzCcKWz2BIqh4a8J+B/DPie2urjxZHdm3ORb+Uw3Pnrn0r3y3vdN8RWZ8qW31C2bBKghx68iqir7g/I+VvBfg69+InjSRNTLWyN/pV3I67DtJHA9M10n7TnwZs9J8C6V4x8KaXGNd8GXKaiqquXurZf9dG3rlefwr3W88F6ddXGp3CoY7i/txbyMpwML93A9q+cfFUnjLwBqUdtqV7cTWpZhE0kheO4ToQR3GDjBo/h9AWpP4Psx8XLzTL+K2XNwBeQXOw+RPbnnaT0yucfhXr/AMXPCGlQfC/xPdWtha295Y6XcXMEwgU7GjjLj8CVxXy98G/j5efBbUNT+H89gJNOupZL3w/PI+FiVzmSD3CsSQM9DXb+KtZ8Q/E74beL7lb6Wzls9LunuWDlYZIxC5xjscZFCstBankPg+w8b3PhXwn4nsteFvqOpWEWpxrHpqSxpnnGM5wCPWt/wmdY+KnxTufA/j2Gw0zWNVs2v7S8s4W+x6pCOJMIxyki85GSOD7Z9y/ZJi8P6h+z/wDDnVjd20tzZ6Kto5MylYyCQ6sOxByOawp9Q0T4jftPaD4g0ee1Twp4D066hudWhKiC4vZ8D7PEw4fYvzNt4BOOpp8i6hc8+8I+LIvgvrPjz4O/aZtU0URC70W6k+9DuIEkJ9VVhxXe/CnwH4e1VorzXNTt3nf5oNNEoVnI/vf4Vw2teH/CPjr9qnVHsvEK2tnaaNunmdCw+0SyZ2A/Tmuk8ZfCPU/Ctr/aFrOmq6Vni6tjnb7kDpUO97jWx2HxP+JXiKzjbRrHR5tAsVXZ5gT5iuSPlI4AryTSND1DxBeR29hbS3U0rY3KpI+pNfVHwpuW8S/DnSpdTRLpirJiZQ3CsVGc98Cuus9Ls9PZja2kNuW+8YowufyFVyc2twTscz8L/Aq+BPDcdo7b7yX57hweN3oPpXYUUVrsSFFFFMAooooAKKKKAKGuaLa+IdNlsL2PzbWXG+PON2DnH6V87/HK18QQ6pFDeXMK6exEdnY2r9V/3O/1r6SubhLW3lmkO2ONS7H2AzXx94g8a3us+Lr/AFQAyXjs0Vt/F5KgkfKPWsqjVio3HR+C7HTbUS6/qyafK3Is7dfNnA/2h0U/Wrmh6toug3iy6Nruq6fcKchpYQYz/vBT0rD8J+G7rxp4kt9NhdvNnfMszZJRe7Gu21rxZo3gG8Ok6N4etb0wHbLfX6EvKw4YrnoMisV3K12PaPh/8Qo/FEK2108P24DIkt2zHMPVe4Psa4349+BNb8S3Vhe6bE15DEuxoV6ofUD3/pXK+GvEmi+LNQhS2t4/C3iNWzby25xbTHPCEdic9a9N8WfFlfCHh+xubq0zqMsnlS2xONmMbz+uR9a2upRsydUz5D8b/C2bxw3/AAj8kclrrUMoNtNGf3kEw6EH+ldJ8K9W+Ifwtj1LRfHvws1nxbYXsQhNxo8CTw3Cg9ZI2Ixn0r1WPwjJD8QB4y8wN4bUf2kt1nrx9zHrmvQ/C3xW/wCEw8M3l3ZWe3U45PIjts5BZvunPYev0qY6bgz568QfE7wjZ2d1Zab+zzqNvftw0dzoUAUe5UH5j7V5fqnj3xD4pjjstH8LXmjMDtE+rrHb21qPVII/T04r6L13xdpXgnUmbyx4k8Rht01zOx+zxMeSEHt61o6Dqnh/4wRvY3+jQ2WqNy89oNrA84fPce1DdwPEvBvwhstF0md9G1xdf1e5bz783KeVcTSY5Kg8EDoAO1erfA201+5v5YLa4hfScmK8sbl87R3IQ8/lXnvibw9d+DfEdxp07Fbm3f8AdyJwWHVWFanhzxxdaL4usdUYGO4VljusDHmrnHI9cVF9dSraH1to+kWuh6fFZWcflW8f3Vznqc1dpkMomhSRTlXUMMe9PrqMwooooAKKKKACiiigAooooAgvrOPULOe2mBMUyFGwcHBGDXM6D8LfDfh2R5bTT0NwwIM0vzNg11tFKyYHmHw8+E9z4K8YalqTzQz2s6MkO3O9QSDgj8K6Hx98PdO8Y6DPbG2iiu1UtBMigMrduR2rrqKXKrWGfDF5azaZfS28waG5t3KkdGVgevtXR+IvHDeLPDlpa6mhOp2TYju1H+uTGMP/ALQwOa96+IXwS03xrfHUIZ20++YfOyAFZPcj1rhtD+G/gW08QNo19f3V/q3IWCVTCpYDoG71g4NFXucP/wAJ1K3wuj8P7+Vuznnny8E/zpvh/wAe/wDCJ+E7mx0yMrqd4x827b/lmvTC+/vXZofBjeIjow8H6l/aQl8kwbxjGcbutTa58OPA154gOj6fqF3Y6mOGhjQzLuPYt2x3os+49Dx2ys7jVr6G1gBlubhwi9ySa+vfAPgOw8F6Hb20UCNdY3TTsMszEc89h7Vg/Dz4Lad4Hvft8kzX98BiN3XAjz1wPWvR60hG2rJk7nm3j/4Sf8Jl4s0rVkuY7eO3KieMrzIoOePfoK29e+FXhrxDIktzpyJOuAJYflPH6V11FXyokhtLVLK1it4gRHGoVcnPAqaiiqAKKKKACiiigAooooAKKKKACiiigAooooAK8C/aOt9Qh1bSLyCJ1gjBKTwrysnuR+Fe+1FcWsN3H5c8STR5ztkUMPyNTKPMrDTsfM7/ABtuW0MA6dGPE2zyDqQUb9n+NdL+zbb3z32t3dxExhl2kyyLyz89DXIW+g3/APwuxbY2eXW88xk8vKeX646Yr6it7eK1jEcMSQxjosahR+QrOKbd2DJaKKK2EFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUARfZohOZ/KTziNvmbRux6ZqWiigAooooAKKKKACiiigAooooAKKKKACiiigD/9k=";

/* ----------------------------- constants ----------------------------- */
const TAX_RATE = 0.09;
const STORE_KEY = "dd_schedule_items_v1";
const STORE_DISC = "dd_retailer_discounts_v1";
const STORE_CLIENTS = "dd_clients_v1";
const itemsKey = (id) => `dd_items_${id}`;
const discKey = (id) => `dd_disc_${id}`;
const contractorsKey = (id) => `dd_contractors_${id}`;
const areasKey = (id) => `dd_areas_${id}`;
const pinboardKey = (id) => `dd_pinboard_${id}`;
const filesKey = (id) => `dd_files_${id}`;
const presKey = (id) => `dd_pres_${id}`;
const notifsKey = (id) => `dd_notifs_${id}`;
const paymentsKey = (id) => `dd_payments_${id}`;
const DEFAULT_AREAS = ["Foyer", "Living Room", "Kitchen", "Dining", "Primary Bed", "Primary Bath", "Office", "Throughout"];
const accessCode = () => Math.random().toString(36).slice(2, 6).toUpperCase() + "-" + Math.random().toString(36).slice(2, 5).toUpperCase();
const mkClient = (o = {}) => ({
  id: o.id || uid(), projectName: o.projectName || "Untitled Project", clientName: o.clientName || "",
  address: o.address || "", status: o.status || "Active", createdAt: o.createdAt || today(),
  coverPhoto: o.coverPhoto || "",
  clientAccess: o.clientAccess ?? false, accessCode: o.accessCode || accessCode(),
});
const TRADES = ["Cabinets", "Painting", "Tile", "Flooring", "Countertops", "Electrical", "Plumbing",
  "Millwork / Carpentry", "Drywall", "Window Treatments", "Wallpaper", "Landscaping", "General Contractor", "Other"];
const CONTRACTOR_STATUS = ["Not started", "In progress", "Complete"];
const CSTATUS_COLOR = {
  "Not started": ["#ECE6DC", "#6B5E4B"],
  "In progress": ["#DEE6EC", "#3E5A72"],
  "Complete": ["#D9E7D9", "#356037"],
};
const mkContractor = (o = {}) => ({
  id: o.id || uid(), company: o.company || "", trade: o.trade || "", scope: o.scope || "", room: o.room || "",
  contractAmount: o.contractAmount ?? "", payments: o.payments || [], status: o.status || "Not started",
  phone: o.phone || "", email: o.email || "", notes: o.notes || "", invoices: o.invoices || [],
});
const paidOf = (c) => (c.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
const balanceOf = (c) => (Number(c.contractAmount) || 0) - paidOf(c);
const payStatus = (c) => {
  const amt = Number(c.contractAmount) || 0, paid = paidOf(c);
  if (amt > 0 && paid >= amt) return "Paid in full";
  if (paid > 0) return "Partially paid";
  return "Unpaid";
};
const PAYSTATUS_COLOR = {
  "Paid in full": ["#D9E7D9", "#356037"],
  "Partially paid": ["#F6E6CC", "#9A6B16"],
  "Unpaid": ["#F2DAD6", "#A33A2C"],
};

const STATUSES = [
  "Idea", "Proposed to client", "Client approved", "Ordered",
  "Shipped", "Delivered", "Installed", "Use existing", "On hold",
];
// statuses a client should see in the read-only preview
const CLIENT_STATUSES = ["Proposed to client", "Client approved", "Ordered", "Shipped", "Delivered", "Installed"];
// statuses where a payment has been made
const ORDERED_STATUSES = ["Ordered", "Shipped", "Delivered", "Installed"];
const PAY_METHODS = ["Credit card", "Check", "Client check", "Wire / ACH", "Zelle", "Cash", "Trade account", "Other"];

// client payments / checks ledger
const PAYMENT_METHODS = ["Check", "Wire / ACH", "Credit card", "Zelle", "Cash", "Other"];
const PAYMENT_CATEGORIES = ["Retainer / Deposit", "Design fee", "Furniture", "Fixtures / Remodel", "Progress payment", "Final payment", "Reimbursement", "Other"];
const PAYMENT_STATUS = ["Pending", "Deposited", "Cleared", "Bounced", "Refunded"];
const PAYMENT_STATUS_COLOR = {
  "Pending": ["#EFEAD6", "#7A6A2E"],
  "Deposited": ["#DEE6EC", "#3E5A72"],
  "Cleared": ["#DCEBDD", "#3C6B3E"],
  "Bounced": ["#F2DDD8", "#9A3B2C"],
  "Refunded": ["#E7E2EC", "#6A5A78"],
};
const mkPayment = (o = {}) => ({
  id: o.id || uid(), date: o.date || today(), method: o.method || "Check", reference: o.reference || "",
  amount: o.amount ?? "", category: o.category || "Progress payment", status: o.status || "Deposited", notes: o.notes || "",
});

const STATUS_COLOR = {
  "Idea": ["#ECE6DC", "#6B5E4B"],
  "Proposed to client": ["#F1DDD0", "#A4502A"],
  "Client approved": ["#E4EBDD", "#4F6B3A"],
  "Ordered": ["#DEE6EC", "#3E5A72"],
  "Shipped": ["#DEE6EC", "#3E5A72"],
  "Delivered": ["#D9E7D9", "#356037"],
  "Installed": ["#CFE0CF", "#27502A"],
  "Use existing": ["#E8E4DC", "#7A7263"],
  "On hold": ["#F6E6CC", "#9A6B16"],
};
const APPROVAL_COLOR = {
  Pending: ["#F6E6CC", "#9A6B16"],
  Approved: ["#D9E7D9", "#356037"],
  Rejected: ["#F2DAD6", "#A33A2C"],
};

const RETAILER_MAP = {
  "wayfair.com": "Wayfair", "westelm.com": "West Elm", "potterybarn.com": "Pottery Barn",
  "livingspaces.com": "Living Spaces", "article.com": "Article", "amazon.com": "Amazon",
  "build.com": "Build.com", "target.com": "Target", "ruggable.com": "Ruggable",
  "rugsusa.com": "Rugs USA", "perigold.com": "Perigold", "birchlane.com": "Birch Lane",
  "roomandboard.com": "Room & Board", "hernest.com": "Hernest", "frontgate.com": "Frontgate",
  "sierralivingconcepts.com": "Sierra Living Concepts", "nathanjames.com": "Nathan James",
  "luxedecor.com": "LuxeDecor", "hulalahome.com": "Hulala Home", "eurekaergonomic.com": "Eureka Ergonomic",
};

const BASE_ROOMS = ["Foyer", "Front Living", "Bar", "Kitchen", "Great Room", "Master Bed",
  "Master Bathroom", "Guest Bed", "Nanny", "Rec. Room", "Gazebo Patio", "Poolside", "Throughout"];
const BASE_TYPES = ["Sofa", "Sectional", "Chair", "Coffee Table", "End Table", "Rug", "Lamp",
  "Floor Lamp", "Nightstand", "Dresser", "Bed", "Dining Table", "Console", "Faucet", "Cabinet Pull", "Light Fixture"];

/* ----------------------------- helpers ----------------------------- */
const uid = () => Math.random().toString(36).slice(2, 9);
const money = (n) =>
  (Number(n) || 0).toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 2 });
const today = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) => (iso ? new Date(iso).toLocaleDateString("en-US", { month: "short", day: "numeric" }) : "—");
const nowTs = () => new Date().toISOString();
const fmtWhen = (iso) => {
  if (!iso) return "";
  const d = new Date(iso), diff = (Date.now() - d.getTime()) / 1000;
  if (diff < 60) return "just now";
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
  if (diff < 604800) return `${Math.floor(diff / 86400)}d ago`;
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
};

/* ---- Excel export (mirrors the original D&D schedule workbook) ---- */
function exportScheduleXlsx(selItems, client, discounts = {}) {
  const TITLE = client?.projectName || "Schedule";
  const HEAD = ["Room / Area", "Item Type", "Source / Retailer", "Product Name", "SKU / Model", "Finish / Selection",
    "Dimensions", "Product URL", "Photo URL", "Price Each", "Disc %", "Net Each", "Qty", "Line Total", "Est. Tax (9%)",
    "Line + Tax", "Availability", "Lead Time / ETA", "Order Status", "Client Approval", "Approval Date", "Actual Cost", "Last Checked", "Notes"];
  const rows = [];
  rows.push([TITLE + " — Furniture & Procurement Schedule"]);
  rows.push(["Exported " + new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" }) + (client?.clientName ? "  ·  " + client.clientName : "")]);
  rows.push([]);
  rows.push(HEAD);

  const byRoom = {};
  selItems.forEach((i) => { (byRoom[i.room || "Unassigned"] ||= []).push(i); });

  const summary = [];
  let gSub = 0;
  Object.entries(byRoom).forEach(([room, list]) => {
    rows.push([room]);
    let sub = 0;
    list.forEach((it) => {
      const price = effPrice(it), disc = discPct(it, discounts), net = netEach(it, discounts);
      const qty = Number(it.qty) || 0, lt = net * qty, tax = lt * TAX_RATE;
      sub += lt;
      rows.push([it.room || "", it.itemType || "", it.retailer || "", it.name || "", it.sku || "", it.finish || "",
        it.dimensions || "", it.url || "", it.imageUrl || "", price || 0, disc || 0, net || 0, qty, lt, tax, lt + tax,
        it.availability || "", it.leadTime || "", it.status || "", it.approval?.state || "", it.approval?.date || "",
        it.actualCost === "" || it.actualCost == null ? "" : Number(it.actualCost), it.lastChecked || "", it.notes || ""]);
    });
    const subtax = sub * TAX_RATE;
    rows.push(["", "", "", "", "", "", "", "", "", "", "", "", `${room} subtotal`, sub, subtax, sub + subtax]);
    rows.push([]);
    summary.push([room, list.length, sub, subtax, sub + subtax]);
    gSub += sub;
  });
  const gtax = gSub * TAX_RATE;
  rows.push(["", "", "", "", "", "", "", "", "", "", "", "", "PROJECT SUBTOTAL", gSub]);
  rows.push(["", "", "", "", "", "", "", "", "", "", "", "", "ESTIMATED TAX (9%)", gtax]);
  rows.push(["", "", "", "", "", "", "", "", "", "", "", "", "PROJECT TOTAL", gSub + gtax]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [16, 16, 18, 32, 14, 22, 20, 32, 30, 11, 7, 11, 5, 13, 13, 13, 14, 16, 18, 15, 13, 12, 13, 34].map((wch) => ({ wch }));
  const money = new Set([9, 11, 13, 14, 15, 21]), pct = new Set([10]);
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell && typeof cell.v === "number") { if (money.has(C)) cell.z = '$#,##0.00'; else if (pct.has(C)) cell.z = '0"%"'; }
    }
  }

  const sRows = [["Room / Area", "Items", "Subtotal", "Est. Tax (9%)", "Total"], ...summary, [], ["PROJECT", selItems.length, gSub, gtax, gSub + gtax]];
  const ws2 = XLSX.utils.aoa_to_sheet(sRows);
  ws2["!cols"] = [{ wch: 22 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  const r2 = XLSX.utils.decode_range(ws2["!ref"]);
  for (let R = r2.s.r; R <= r2.e.r; R++) for (let C = 2; C <= 4; C++) { const c = ws2[XLSX.utils.encode_cell({ r: R, c: C })]; if (c && typeof c.v === "number") c.z = '$#,##0.00'; }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Schedule");
  XLSX.utils.book_append_sheet(wb, ws2, "Summary");
  XLSX.writeFile(wb, (TITLE.replace(/[^\w]+/g, "_").replace(/^_|_$/g, "") || "Project") + "_Schedule.xlsx");
}
// read uploaded files into {id,name,type,size,dataUrl} attachments (cap ~4MB for browser storage)
async function readAttachments(fileList, extra = {}) {
  const files = Array.from(fileList || []);
  return (await Promise.all(files.map((f) => new Promise((res) => {
    if (f.size > 4 * 1024 * 1024) { res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: "", tooBig: true, addedAt: today(), ...extra }); return; }
    const r = new FileReader();
    r.onload = () => res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: r.result, addedAt: today(), ...extra });
    r.onerror = () => res(null);
    r.readAsDataURL(f);
  })))).filter(Boolean);
}

const effPrice = (it) =>
  it.priceOverride !== null && it.priceOverride !== "" && !isNaN(it.priceOverride)
    ? Number(it.priceOverride)
    : Number(it.priceAuto) || 0;
// discount %: per-item override if set, else the retailer's default
const discPct = (it, rd = {}) =>
  it.discountPct !== "" && it.discountPct !== null && it.discountPct !== undefined && !isNaN(it.discountPct)
    ? Number(it.discountPct)
    : Number(rd[it.retailer]) || 0;
// net (after-discount) price each — her actual cost
const netEach = (it, rd = {}) => effPrice(it) * (1 - (discPct(it, rd) || 0) / 100);
// line total: useNet=true → her cost after discount; false → retail (what a client sees)
const lineTotal = (it, rd = {}, useNet = true) =>
  (Number(it.qty) || 0) * (useNet ? netEach(it, rd) : effPrice(it));

function retailerFromUrl(url) {
  try {
    const host = new URL(url).hostname.replace(/^www\./, "");
    if (RETAILER_MAP[host]) return RETAILER_MAP[host];
    const base = host.split(".").slice(-2, -1)[0] || host;
    return base.charAt(0).toUpperCase() + base.slice(1);
  } catch {
    return "";
  }
}

/* ----------------------------- seed data ----------------------------- */
const seed = () => [
  mk({ room: "Foyer", itemType: "Console", retailer: "Danbury House", name: "Use Existing", status: "Use existing", qty: 1, priceAuto: 0 }),
  mk({ room: "Front Living", itemType: "Sofa", retailer: "Living Spaces", name: 'Nomad Foam 90" Ivory Performance Fabric Estate Sofa', sku: "330391", finish: "Ivory Performance Fabric", url: "https://www.livingspaces.com/pdp-nomad-90-inch-estate-sofa-330391", dimensions: '90"W × 40"D × 38"H', priceAuto: 1295, qty: 1, status: "Proposed to client", availability: "In stock" }),
  mk({ room: "Front Living", itemType: "Coffee Table", retailer: "Hernest", name: 'Marisol 55" Oak Coffee Table', sku: "8522", finish: "Warm Brown Oak", url: "https://www.hernest.com/products/marisol-55-inch-oak-coffee-table-p-8522.html", priceAuto: 949, qty: 1, status: "Client approved", approval: { state: "Approved", date: today() }, availability: "In stock" }),
  mk({ room: "Front Living", itemType: "End Table", retailer: "Wayfair", name: "Joss & Main Sumiya Antique Brass Accent Table", sku: "W112145866", finish: "Antique Brass / Travertine Top", url: "https://www.wayfair.com/furniture/pdp/joss-main-sumiya-antique-brass-accent-table-w112145866.html", priceAuto: 432.99, qty: 2, status: "Proposed to client", availability: "In stock" }),
  mk({ room: "Front Living", itemType: "Rug", retailer: "Ruggable", name: "Haze Grey Tufted Rug 9'×12'", finish: "Grey / 9×12", url: "https://ruggable.com/products/haze-grey-tufted-rug", dimensions: "9' × 12'", priceAuto: 999, qty: 1, status: "Idea", availability: "Out of stock" }),
  mk({ room: "Kitchen", itemType: "Counterstool", retailer: "Wayfair", name: "Corrigan Studio Riccio Vegan Leather Upholstered Stool", sku: "W114978162", finish: "Beige / Brushed Brass Base", url: "https://www.wayfair.com/furniture/pdp/corrigan-studio-riccio-upholstered-stool-w114978162.html", priceAuto: 409.99, qty: 3, status: "On hold", availability: "In stock" }),
  mk({ room: "Kitchen", itemType: "Dining Table", retailer: "Pottery Barn", name: 'Modern Farmhouse Round Pedestal Extending Dining Table (60"–78")', sku: "SPAF", finish: "Tahoe Brown", url: "https://www.potterybarn.com/products/modern-farmhouse-round-pedestal-extending-dining-table/", priceAuto: 1999.2, qty: 1, status: "Ordered", approval: { state: "Approved", date: today() }, leadTime: "6–8 wks", availability: "In stock" }),
  mk({ room: "Kitchen", itemType: "Chair", retailer: "West Elm", name: "Adler Dining Chair", sku: "H12188", finish: "Slate Performance Velvet", url: "https://www.westelm.com/products/adler-dining-chair-h12188/", priceAuto: 279.2, qty: 6, status: "Client approved", approval: { state: "Approved", date: today() }, availability: "In stock" }),
  mk({ room: "Great Room", itemType: "Sectional", retailer: "Living Spaces", name: 'Cambrie 124" 2-Piece Dual Chaise Sectional', sku: "266019", finish: "Fuzzy White Fabric", url: "https://www.livingspaces.com/pdp-cambrie-2-piece-sectional-266019", dimensions: '124"W', priceAuto: 1395, qty: 1, status: "Delivered", approval: { state: "Approved", date: today() }, actualCost: 1395, availability: "In stock" }),
  mk({ room: "Great Room", itemType: "Floor Lamp", retailer: "Article", name: "Zebu Small Linen Floor Lamp – Brass", sku: "27462", finish: "Ecru White Linen / Antique Brass", url: "https://www.article.com/product/27462/zebu-small-linen-floor-lamp-brass", priceAuto: 249, qty: 1, status: "Shipped", approval: { state: "Approved", date: today() }, leadTime: "Ships in 1 wk", availability: "In stock" }),
  mk({ room: "Master Bed", itemType: "Bed", retailer: "Living Spaces", name: "Porto King Upholstered Platform Bed (Berkus + Brent)", sku: "314825", finish: "Natural Fabric", url: "https://www.livingspaces.com/pdp-porto-eastern-king-upholstered-platform-bed-314825", priceAuto: 1595, qty: 1, status: "Proposed to client", availability: "In stock", lastPriceChange: { old: 1395, new: 1595, date: today() } }),
  mk({ room: "Master Bed", itemType: "Nightstand", retailer: "West Elm", name: 'Clementine Marble Nightstand (22")', sku: "H14127", finish: "Volakas Marble / Champagne Bronze", url: "https://www.westelm.com/products/clementine-nightstand-h14127/", priceAuto: 720, qty: 2, status: "Idea", availability: "In stock" }),
  mk({ room: "Rec. Room", itemType: "Rug", retailer: "Wayfair", name: "Canora Grey Rockett Modern Abstract Blue Area Rug", sku: "W115591766", finish: "Blue / 10' × 12'8\"", url: "https://www.wayfair.com/rugs/pdp/canora-grey-rockett-modern-abstract-blue-area-rug-w115591766.html", priceAuto: 1762, qty: 1, status: "Proposed to client", approval: { state: "Rejected", date: today() }, notes: "Client felt too bold — sourcing alternative", availability: "In stock" }),
  mk({ room: "Gazebo Patio", itemType: "Bar Cabinet", retailer: "Sierra Living Concepts", name: "Clewiston Outdoor Bar Cabinet with Mini Fridge Space", sku: "329777165IND", finish: "Natural Honey", url: "https://www.sierralivingconcepts.com/product/14885/clewiston-outdoor-bar-cabinet-with-fridge-space", priceAuto: 3149, qty: 1, status: "Idea", availability: "In stock" }),
  mk({ room: "Poolside", itemType: "Chaise Lounge", retailer: "West Elm", name: "Porto Aluminum Outdoor Chaise Lounge", finish: "Dark Bronze", url: "https://www.westelm.com/products/porto-aluminum-outdoor-chaise-lounger-f344/", priceAuto: 1440, qty: 4, status: "Proposed to client", availability: "In stock" }),
];

function mk(o) {
  return {
    id: uid(), project: "D & D Mercer",
    room: o.room || "", itemType: o.itemType || "", retailer: o.retailer || "",
    name: o.name || "", sku: o.sku || "", finish: o.finish || "", url: o.url || "",
    imageUrl: o.imageUrl || "", dimensions: o.dimensions || "",
    priceAuto: o.priceAuto ?? 0, priceOverride: o.priceOverride ?? "", discountPct: o.discountPct ?? "",
    qty: o.qty ?? 1, availability: o.availability || "Unknown", leadTime: o.leadTime || "",
    status: o.status || "Idea",
    approval: o.approval || { state: "Pending", date: "" },
    actualCost: o.actualCost ?? "", lastChecked: o.lastChecked || today(),
    notes: o.notes || "", lastPriceChange: o.lastPriceChange || null,
    orderNumber: o.orderNumber || "", carrier: o.carrier || "", trackingNumber: o.trackingNumber || "",
    trackingUrl: o.trackingUrl || "", shipStatus: o.shipStatus || "", shipChecked: o.shipChecked || "",
    orderDate: o.orderDate || "", eta: o.eta || "", lastScanned: o.lastScanned || "", scanResult: o.scanResult || "",
    payMethod: o.payMethod || "", payRef: o.payRef || "",
    receipts: o.receipts || [],
  };
}

/* --------------------- live product fetch via Claude --------------------- */
async function fetchProductData(url) {
  const prompt =
    `Look up this exact product page and extract its current details. URL: ${url}\n\n` +
    `Use web search to open the retailer's product page. Respond with ONLY a JSON object (no markdown, no commentary) with EXACTLY these keys:\n` +
    `{"name": string, "sku": string, "price": number or null (USD digits only), "dimensions": string, ` +
    `"availability": "In stock" | "Out of stock" | "Unknown", "imageUrl": string, "retailer": string}\n\n` +
    `For "imageUrl": return the listing's MAIN product photo as a direct, hotlinkable HTTPS image URL taken from the product page itself — ` +
    `prefer the page's Open Graph image (og:image meta tag) or the primary gallery image hosted on the retailer's own image CDN ` +
    `(a URL that typically ends in .jpg, .jpeg, .png or .webp, or contains the retailer's media/image host). ` +
    `Do NOT return a search-results thumbnail, a Google/Bing image link, a logo, or a page URL. If you cannot find a real product image URL, use "".\n` +
    `If any other field is unknown use "" or null. Do not invent values.`;
  try {
    const res = await fetch("/api/anthropic", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 1000,
        messages: [{ role: "user", content: prompt }],
        tools: [{ type: "web_search_20250305", name: "web_search" }],
      }),
    });
    const data = await res.json();
    const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a === -1 || b === -1) return null;
    return JSON.parse(text.slice(a, b + 1));
  } catch (e) {
    console.error("fetch product failed", e);
    return null;
  }
}

// realistic starting trade discounts so the recalc is visible on load — she edits these
const seedDiscounts = () => ({ "West Elm": 20, "Pottery Barn": 20, "Living Spaces": 10, "Room & Board": 10 });

const seedContractors = () => [
  mkContractor({ company: "Obe — Custom Cabinetry", trade: "Cabinets", scope: "Bar cabinets + Great Room media console (custom)", room: "Bar", contractAmount: 7500, payments: [{ id: uid(), amount: 3750, date: today(), note: "50% deposit" }], status: "In progress", phone: "(602) 555-0142" }),
  mkContractor({ company: "Omar — Stone & Surfaces", trade: "Countertops", scope: "Bar top fabrication & install (labor)", room: "Bar", contractAmount: 2300, payments: [{ id: uid(), amount: 2300, date: today(), note: "Paid in full" }], status: "Complete" }),
  mkContractor({ company: "Desert Sky Painting", trade: "Painting", scope: "Whole-house interior paint", room: "Throughout", contractAmount: 6500, payments: [{ id: uid(), amount: 2000, date: today(), note: "Deposit" }], status: "In progress", email: "jobs@desertskypaint.com" }),
];
// areas for the seeded project, in the order they appear in her schedule
const seedAreas = () => {
  const order = [];
  seed().forEach((i) => { if (i.room && !order.includes(i.room)) order.push(i.room); });
  return order;
};

/* --------------------- inbox order lookup via Gmail (MCP) --------------------- */
async function fetchOrderInfo(item, client) {
  const prompt =
    `Search my Gmail for order confirmation, receipt, and shipping/tracking emails about this furniture order.\n` +
    `Retailer: ${item.retailer || "unknown"}\nProduct: ${item.name || ""}\nSKU/Model: ${item.sku || ""}\n` +
    (client?.clientName ? `Client/recipient may be: ${client.clientName}\n` : "") +
    `Look for the most recent matching email. Respond with ONLY a JSON object (no markdown, no commentary):\n` +
    `{"found": boolean, "orderPlaced": boolean, "orderNumber": string, "orderDate": "YYYY-MM-DD" or "", ` +
    `"carrier": string, "trackingNumber": string, "estimatedDelivery": "YYYY-MM-DD" or "", ` +
    `"statusGuess": "Ordered" | "Shipped" | "Delivered" | "Unknown"}\n` +
    `If no clearly matching email exists, return {"found": false}. Do not invent tracking numbers.`;
  try {
    const res = await fetch("/api/anthropic", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 1200,
        messages: [{ role: "user", content: prompt }],
        mcp_servers: [{ type: "url", url: "https://gmailmcp.googleapis.com/mcp/v1", name: "gmail" }],
      }),
    });
    const data = await res.json();
    const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a === -1 || b === -1) return { found: false };
    return JSON.parse(text.slice(a, b + 1));
  } catch (e) {
    console.error("inbox scan failed", e);
    return null;
  }
}

/* --------------------- shipment status from tracking # / URL --------------------- */
async function fetchShipmentStatus(item) {
  const ref = [item.carrier, item.trackingNumber].filter(Boolean).join(" ") || item.trackingUrl;
  if (!ref) return null;
  const prompt =
    `Find the current shipping status for this package. ${item.carrier ? "Carrier: " + item.carrier + ". " : ""}` +
    `${item.trackingNumber ? "Tracking number: " + item.trackingNumber + ". " : ""}` +
    `${item.trackingUrl ? "Tracking URL: " + item.trackingUrl + ". " : ""}\n` +
    `Use web search to find the latest tracking status. Respond with ONLY a JSON object (no markdown):\n` +
    `{"found": boolean, "status": string (e.g. "In transit", "Out for delivery", "Delivered"), ` +
    `"location": string, "estimatedDelivery": "YYYY-MM-DD" or "", "delivered": boolean, "lastUpdate": string}\n` +
    `If status cannot be determined, return {"found": false}. Do not guess.`;
  try {
    const res = await fetch("/api/anthropic", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 800,
        messages: [{ role: "user", content: prompt }],
        tools: [{ type: "web_search_20250305", name: "web_search" }],
      }),
    });
    const data = await res.json();
    const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a === -1 || b === -1) return { found: false };
    return JSON.parse(text.slice(a, b + 1));
  } catch (e) {
    console.error("shipment status failed", e);
    return null;
  }
}

/* ----------------------------- storage ----------------------------- */
/* store provided by ./lib/store (Supabase) */

/* ============================== APP ============================== */
/* ============================== APP (clients + routing) ============================== */
export default function App() {
  const [clients, setClients] = useState([]);
  const [data, setData] = useState({});        // { [clientId]: { items, discounts } }
  const [currentId, setCurrentId] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [authUser, setAuthUser] = useState(undefined); // undefined=checking, null=logged out, obj=signed in
  const [session, setSession] = useState({ role: "designer" }); // role view (designer | client preview)
  const [portal, setPortal] = useState(null); // external client portal (code + read-only payload)

  // real auth gate (Supabase)
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setAuthUser(data.session?.user || null));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setAuthUser(s?.user || null));
    return () => sub.subscription.unsubscribe();
  }, []);

  // load this designer's data once signed in
  useEffect(() => {
    if (!authUser) { setLoaded(false); setClients([]); setData({}); setCurrentId(null); setSession({ role: "designer" }); return; }
    (async () => {
      let cl = await store.load(STORE_CLIENTS);
      const d = {};
      if (!cl || !cl.length) {
        const def = mkClient({ id: "mercer", projectName: "D & D Mercer", clientName: "Mercer Residence", address: "New House", status: "Active", clientAccess: true, accessCode: "MERCER-25" });
        cl = [def];
        d[def.id] = { items: seed(), discounts: seedDiscounts(), contractors: seedContractors(), areas: seedAreas(), pinboard: [], files: [], presentations: [], notifications: [], payments: [
          mkPayment({ method: "Check", reference: "#1042", amount: 15000, category: "Retainer / Deposit", status: "Cleared", notes: "Initial retainer" }),
          mkPayment({ method: "Check", reference: "#1061", amount: 22000, category: "Furniture", status: "Deposited", notes: "First furniture milestone" }),
        ] };
        await store.save(STORE_CLIENTS, cl);
        await store.save(itemsKey(def.id), d[def.id].items);
        await store.save(discKey(def.id), d[def.id].discounts);
        await store.save(contractorsKey(def.id), d[def.id].contractors);
        await store.save(areasKey(def.id), d[def.id].areas);
      } else {
        for (const c of cl) {
          d[c.id] = {
            items: (await store.load(itemsKey(c.id))) || [],
            discounts: (await store.load(discKey(c.id))) || {},
            contractors: (await store.load(contractorsKey(c.id))) || [],
            areas: (await store.load(areasKey(c.id))) || DEFAULT_AREAS,
            pinboard: (await store.load(pinboardKey(c.id))) || [],
            files: (await store.load(filesKey(c.id))) || [],
            presentations: (await store.load(presKey(c.id))) || [],
            notifications: (await store.load(notifsKey(c.id))) || [],
            payments: (await store.load(paymentsKey(c.id))) || [],
          };
        }
      }
      setClients(cl); setData(d); setLoaded(true);
    })();
  }, [authUser]);

  const setClientItems = (id, updater) =>
    setData((prev) => {
      const prevItems = prev[id]?.items || [];
      const items = typeof updater === "function" ? updater(prevItems) : updater;
      store.save(itemsKey(id), items);
      return { ...prev, [id]: { ...prev[id], items } };
    });
  const setClientDisc = (id, updater) =>
    setData((prev) => {
      const prevDisc = prev[id]?.discounts || {};
      const discounts = typeof updater === "function" ? updater(prevDisc) : updater;
      store.save(discKey(id), discounts);
      return { ...prev, [id]: { ...prev[id], discounts } };
    });

  const setClientContractors = (id, updater) =>
    setData((prev) => {
      const prevC = prev[id]?.contractors || [];
      const contractors = typeof updater === "function" ? updater(prevC) : updater;
      store.save(contractorsKey(id), contractors);
      return { ...prev, [id]: { ...prev[id], contractors } };
    });
  const setClientAreas = (id, updater) =>
    setData((prev) => {
      const prevA = prev[id]?.areas || [];
      const areas = typeof updater === "function" ? updater(prevA) : updater;
      store.save(areasKey(id), areas);
      return { ...prev, [id]: { ...prev[id], areas } };
    });
  const setClientPinboard = (id, updater) =>
    setData((prev) => {
      const prevP = prev[id]?.pinboard || [];
      const pinboard = typeof updater === "function" ? updater(prevP) : updater;
      store.save(pinboardKey(id), pinboard);
      return { ...prev, [id]: { ...prev[id], pinboard } };
    });
  const setClientFiles = (id, updater) =>
    setData((prev) => {
      const prevF = prev[id]?.files || [];
      const files = typeof updater === "function" ? updater(prevF) : updater;
      store.save(filesKey(id), files);
      return { ...prev, [id]: { ...prev[id], files } };
    });
  const setClientPres = (id, updater) =>
    setData((prev) => {
      const prevP = prev[id]?.presentations || [];
      const presentations = typeof updater === "function" ? updater(prevP) : updater;
      store.save(presKey(id), presentations);
      return { ...prev, [id]: { ...prev[id], presentations } };
    });
  const setClientNotifs = (id, updater) =>
    setData((prev) => {
      const prevN = prev[id]?.notifications || [];
      const notifications = typeof updater === "function" ? updater(prevN) : updater;
      store.save(notifsKey(id), notifications);
      return { ...prev, [id]: { ...prev[id], notifications } };
    });
  const setClientPayments = (id, updater) =>
    setData((prev) => {
      const prevP = prev[id]?.payments || [];
      const payments = typeof updater === "function" ? updater(prevP) : updater;
      store.save(paymentsKey(id), payments);
      return { ...prev, [id]: { ...prev[id], payments } };
    });

  const addClient = (meta) => {
    const c = mkClient(meta);
    setClients((prev) => { const next = [...prev, c]; store.save(STORE_CLIENTS, next); return next; });
    setData((prev) => ({ ...prev, [c.id]: { items: [], discounts: {}, contractors: [], areas: DEFAULT_AREAS, pinboard: [], files: [], presentations: [], notifications: [], payments: [] } }));
    store.save(itemsKey(c.id), []); store.save(discKey(c.id), {}); store.save(contractorsKey(c.id), []); store.save(areasKey(c.id), DEFAULT_AREAS);
    store.save(pinboardKey(c.id), []); store.save(filesKey(c.id), []); store.save(presKey(c.id), []); store.save(notifsKey(c.id), []); store.save(paymentsKey(c.id), []);
    setCurrentId(c.id);
  };
  const deleteClient = (id) => {
    setClients((prev) => { const next = prev.filter((c) => c.id !== id); store.save(STORE_CLIENTS, next); return next; });
    setData((prev) => { const n = { ...prev }; delete n[id]; return n; });
  };
  const updateClient = (id, patch) =>
    setClients((prev) => { const next = prev.map((c) => (c.id === id ? { ...c, ...patch } : c)); store.save(STORE_CLIENTS, next); return next; });

  const startSession = (s) => setSession(s);
  const signOut = async () => { setCurrentId(null); await supabase.auth.signOut(); };
  const backToDesigner = () => setSession({ role: "designer" });

  // ---- still checking auth ----
  if (authUser === undefined) return <div className="app" style={{ background: "var(--paper)", minHeight: "100vh", padding: 48 }}><Style /><Loader2 className="spin" /></div>;

  // ---- not signed in: branded login, or an external client viewing via access code ----
  if (!authUser) {
    if (portal) {
      const p = portal, noop = () => {};
      return (
        <Schedule client={p.project} items={p.items || []} discounts={p.discounts || {}}
                  contractors={[]} areas={p.areas && p.areas.length ? p.areas : DEFAULT_AREAS}
                  pinboard={p.pinboard || []} files={p.files || []}
                  presentations={p.presentations || []} notifications={p.notifications || []}
                  onItems={noop} onDiscounts={noop} onContractors={noop} onAreas={noop}
                  onPinboard={noop} onFiles={noop} onPres={noop} onNotifs={noop}
                  clientMode portalReadOnly onSignOut={() => setPortal(null)} onBack={() => setPortal(null)} />
      );
    }
    return <Login onPortal={(payload, code) => setPortal({ ...payload, code })} />;
  }

  if (!loaded) return <div className="app" style={{ background: "var(--paper)", minHeight: "100vh", padding: 48 }}><Style /><Loader2 className="spin" /></div>;

  // ---- client portal PREVIEW (designer viewing the client's read-only portal) ----
  if (session.role === "client") {
    const c = clients.find((x) => x.id === session.clientId);
    if (!c || !data[c.id]) { backToDesigner(); return null; }
    const cid = c.id;
    return (
      <Schedule client={c} items={data[cid].items} discounts={data[cid].discounts}
                contractors={data[cid].contractors || []} areas={data[cid].areas || DEFAULT_AREAS}
                pinboard={data[cid].pinboard || []} files={data[cid].files || []}
                presentations={data[cid].presentations || []} notifications={data[cid].notifications || []}
                onItems={(u) => setClientItems(cid, u)} onDiscounts={(u) => setClientDisc(cid, u)}
                onContractors={(u) => setClientContractors(cid, u)} onAreas={(u) => setClientAreas(cid, u)}
                onPinboard={(u) => setClientPinboard(cid, u)} onFiles={(u) => setClientFiles(cid, u)}
                onPres={(u) => setClientPres(cid, u)} onNotifs={(u) => setClientNotifs(cid, u)}
                clientMode onSignOut={backToDesigner} onBack={backToDesigner} />
    );
  }

  // ---- designer workspace ----
  const current = currentId ? clients.find((c) => c.id === currentId) : null;
  if (current && data[currentId]) {
    return (
      <Schedule client={current} items={data[currentId].items} discounts={data[currentId].discounts}
                contractors={data[currentId].contractors || []} areas={data[currentId].areas || DEFAULT_AREAS}
                pinboard={data[currentId].pinboard || []} files={data[currentId].files || []}
                presentations={data[currentId].presentations || []} notifications={data[currentId].notifications || []}
                payments={data[currentId].payments || []} onPayments={(u) => setClientPayments(currentId, u)}
                onItems={(u) => setClientItems(currentId, u)} onDiscounts={(u) => setClientDisc(currentId, u)}
                onContractors={(u) => setClientContractors(currentId, u)} onAreas={(u) => setClientAreas(currentId, u)}
                onPinboard={(u) => setClientPinboard(currentId, u)} onFiles={(u) => setClientFiles(currentId, u)}
                onPres={(u) => setClientPres(currentId, u)} onNotifs={(u) => setClientNotifs(currentId, u)}
                onUpdateClient={(patch) => updateClient(currentId, patch)}
                onPreviewClientLogin={() => startSession({ role: "client", clientId: currentId })}
                onBack={() => setCurrentId(null)} />
    );
  }
  return <Dashboard clients={clients} data={data} onOpen={setCurrentId} onAdd={addClient} onDelete={deleteClient} onSignOut={signOut} />;
}

/* ----------------------------- Login (Supabase auth) ----------------------------- */
function Login({ onPortal }) {
  const [view, setView] = useState("designer"); // designer | client
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [mode, setMode] = useState("signin"); // signin | signup
  const [err, setErr] = useState("");
  const [note, setNote] = useState("");
  const [busy, setBusy] = useState(false);
  const [code, setCode] = useState("");

  const submit = async () => {
    if (!email.trim() || !pw) return;
    setBusy(true); setErr(""); setNote("");
    try {
      if (mode === "signup") {
        const { data, error } = await supabase.auth.signUp({ email: email.trim(), password: pw });
        if (error) throw error;
        if (!data.session) { setNote("Account created. Check your email to confirm, then sign in."); setMode("signin"); }
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pw });
        if (error) throw error;
      }
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };

  const clientEnter = async () => {
    if (!code.trim()) return;
    setBusy(true); setErr("");
    try {
      const { data, error } = await supabase.rpc("portal_get", { p_code: code.trim() });
      if (error) throw error;
      if (!data) { setErr("That access code didn’t match a project. Check the code with your designer."); return; }
      onPortal(data, code.trim());
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };

  return (
    <div className="app login-app"><Style />
      <div className="login-card">
        <div className="login-brand">
          <div className="login-logo"><img src={LOGO} alt="Narciso Designs" /></div>
          <div><div className="login-studio font-display">Narciso Designs</div><div className="login-tag">Project workspace</div></div>
        </div>

        {view === "designer" ? (
          <div className="login-pane">
            <p className="login-copy">{mode === "signup"
              ? "Create your studio account to manage projects, schedules, and client presentations."
              : "Sign in to your studio workspace."}</p>
            <input className="input login-input" type="email" placeholder="Email" value={email}
                   onChange={(e) => { setEmail(e.target.value); setErr(""); }} autoFocus />
            <input className="input login-input" type="password" placeholder="Password" value={pw}
                   onChange={(e) => { setPw(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} />
            {err && <div className="login-err">{err}</div>}
            {note && <div className="login-notice">{note}</div>}
            <button className="btn btn-primary login-btn" onClick={submit} disabled={busy || !email.trim() || !pw}>
              {busy ? "…" : mode === "signup" ? "Create account" : "Sign in"}
            </button>
            <div className="login-switch">
              {mode === "signup" ? "Already have an account? " : "New here? "}
              <span onClick={() => { setMode(mode === "signup" ? "signin" : "signup"); setErr(""); setNote(""); }}>
                {mode === "signup" ? "Sign in" : "Create one"}
              </span>
            </div>
            <button className="login-alt" onClick={() => { setView("client"); setErr(""); }}><KeyRound size={14} /> I’m a client — enter access code</button>
          </div>
        ) : (
          <div className="login-pane">
            <p className="login-copy">Enter the access code your designer shared to view your project — selections, presentations, and shared files.</p>
            <input className="input login-input" placeholder="Access code (e.g. MERCER-25)" value={code}
                   onChange={(e) => { setCode(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && clientEnter()} autoFocus />
            {err && <div className="login-err">{err}</div>}
            <button className="btn btn-primary login-btn" onClick={clientEnter} disabled={busy || !code.trim()}>{busy ? "…" : "View my project"}</button>
            <button className="login-alt" onClick={() => { setView("designer"); setErr(""); }}>← Designer sign-in</button>
          </div>
        )}
        <div className="login-foot">Your account and project data are private to you.</div>
      </div>
    </div>
  );
}

/* ----------------------------- Dashboard ----------------------------- */
function Dashboard({ clients, data, onOpen, onAdd, onDelete, onSignOut }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState({ projectName: "", clientName: "", address: "", coverPhoto: "" });

  const statsFor = (id) => {
    const items = data[id]?.items || [];
    const disc = data[id]?.discounts || {};
    const contractors = data[id]?.contractors || [];
    const total = items.reduce((s, i) => s + lineTotal(i, disc, true) * (1 + TAX_RATE), 0);
    const awaiting = items.filter((i) => i.status === "Proposed to client" && i.approval.state === "Pending").length;
    const inTransit = items.filter((i) => ["Ordered", "Shipped"].includes(i.status)).length;
    const done = items.filter((i) => ["Delivered", "Installed"].includes(i.status)).length;
    const owed = contractors.reduce((s, c) => s + Math.max(0, balanceOf(c)), 0);
    return { count: items.length, total, awaiting, inTransit, done, trades: contractors.length, owed };
  };

  const submit = () => {
    if (!form.projectName.trim() && !form.clientName.trim()) return;
    onAdd({ projectName: form.projectName.trim() || form.clientName.trim(), clientName: form.clientName.trim(), address: form.address.trim(), coverPhoto: form.coverPhoto });
    setForm({ projectName: "", clientName: "", address: "", coverPhoto: "" }); setAdding(false);
  };

  return (
    <div className="app dash-app" style={{ minHeight: "100vh" }}>
      <Style />
      <header className="hdr dash-hdr">
        <img className="hdr-damask" src={LOGO} alt="" aria-hidden="true" />
        <div className="hdr-row">
          <div className="dash-brand">
            <div className="dash-logo"><img src={LOGO} alt="Narciso Designs" /></div>
            <div>
              <div className="eyebrow">Narciso Designs</div>
              <h1 className="font-display title">{"{ Your Projects }"}</h1>
            </div>
          </div>
          <div className="hdr-actions">
            <button className="btn btn-primary" onClick={() => setAdding((v) => !v)}><Plus size={15} /> New project</button>
            <button className="btn btn-ghost" onClick={onSignOut} title="Sign out"><LogOut size={15} /> Sign out</button>
          </div>
        </div>
        {adding && (
          <div className="newclient">
            <div className="newclient-grid">
              <label className="field"><span>Project name</span><input className="input" placeholder="e.g. D & D Mercer" value={form.projectName} onChange={(e) => setForm({ ...form, projectName: e.target.value })} /></label>
              <label className="field"><span>Client</span><input className="input" placeholder="e.g. Mercer Residence" value={form.clientName} onChange={(e) => setForm({ ...form, clientName: e.target.value })} /></label>
              <label className="field"><span>Property / address</span><input className="input" placeholder="Optional" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></label>
            </div>
            <div className="field" style={{ marginTop: 12 }}><span style={{ fontSize: 12, letterSpacing: ".05em", textTransform: "uppercase", color: "var(--ink-soft)", fontWeight: 600, marginBottom: 6, display: "block" }}>Cover photo (the house)</span>
              <ImageInput value={form.coverPhoto} onChange={(v) => setForm({ ...form, coverPhoto: v })} />
            </div>
            <div className="newclient-foot">
              <button className="btn btn-ghost" onClick={() => setAdding(false)}>Cancel</button>
              <button className="btn btn-primary" onClick={submit}><Check size={15} /> Create project</button>
            </div>
          </div>
        )}
      </header>

      <div className="dash">
        {clients.length === 0 && <div className="empty"><Users size={28} /><p>No projects yet. Create your first one above.</p></div>}
        <div className="client-grid">
          {clients.map((c) => {
            const s = statsFor(c.id);
            const unread = (data[c.id]?.notifications || []).filter((n) => n.audience === "designer" && !n.read).length;
            return (
              <div key={c.id} className={"client-card" + (c.coverPhoto ? " has-cover" : "")} onClick={() => onOpen(c.id)}>
                {c.coverPhoto && <div className="client-cover"><img src={c.coverPhoto} alt="" /></div>}
                <div className="client-card-top">
                  <div className="client-avatar font-display">{(c.clientName || c.projectName || "?").charAt(0)}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    {unread > 0 && <span className="card-notif" title={`${unread} new update${unread === 1 ? "" : "s"}`}><Bell size={12} /> {unread}</span>}
                    <button className="icon-btn del" title="Delete project" onClick={(e) => { e.stopPropagation(); if (confirm(`Delete "${c.projectName}" and its schedule?`)) onDelete(c.id); }}><Trash2 size={13} /></button>
                  </div>
                </div>
                <div className="client-name font-display">{c.projectName}</div>
                {c.clientName && <div className="client-sub"><Building2 size={12} /> {c.clientName}</div>}
                {c.address && <div className="client-sub"><MapPin size={12} /> {c.address}</div>}
                <div className="client-total">{money(s.total)} <em>est. total</em></div>
                <div className="client-stats">
                  <span>{s.count} items</span>
                  {s.trades > 0 && <span>{s.trades} contractors</span>}
                  {s.awaiting > 0 && <span className="st-await">{s.awaiting} awaiting approval</span>}
                  {s.inTransit > 0 && <span className="st-transit">{s.inTransit} in transit</span>}
                  {s.done > 0 && <span className="st-done">{s.done} delivered</span>}
                  {s.owed > 0 && <span className="st-await">{money(s.owed)} owed to trades</span>}
                </div>
                <div className="client-open">Open schedule →</div>
              </div>
            );
          })}
        </div>
      </div>
      <footer className="foot">Each project keeps its own schedule, retailer discounts, order tracking, and receipts — all saved to this workspace.</footer>
    </div>
  );
}

function Schedule({ client, items, onItems, discounts, onDiscounts, contractors, onContractors, areas, onAreas, pinboard, onPinboard, files, onFiles, presentations, onPres, notifications = [], onNotifs, payments = [], onPayments, clientMode = false, portalReadOnly = false, onSignOut, onUpdateClient, onPreviewClientLogin, onBack }) {
  const setItems = (u) => onItems(u);
  const setRetailerDiscounts = (u) => onDiscounts(u);
  const retailerDiscounts = discounts;
  const [section, setSection] = useState("overview"); // overview | furnishings | contractors | pinboard | files
  const [clientView, setClientView] = useState(clientMode);
  const [showAccess, setShowAccess] = useState(false);
  const [preview, setPreview] = useState(null); // image src for lightbox
  const [view, setView] = useState(clientMode ? "cards" : "list");
  const [search, setSearch] = useState("");
  const [fRoom, setFRoom] = useState("All");
  const [fStatus, setFStatus] = useState("All");
  const [collapsed, setCollapsed] = useState({});
  const [draft, setDraft] = useState(null);     // add/review panel
  const [urlInput, setUrlInput] = useState("");
  const [fetching, setFetching] = useState(false);
  const [refreshing, setRefreshing] = useState(null); // item id
  const [showDiscounts, setShowDiscounts] = useState(false);
  const [showAreas, setShowAreas] = useState(false);
  const [scanning, setScanning] = useState(null);     // item id being inbox-scanned
  const [scanAll, setScanAll] = useState(false);
  const [tracking, setTracking] = useState(null);      // item id being shipment-tracked
  const [selected, setSelected] = useState(() => new Set()); // selected item ids for export
  const [exporting, setExporting] = useState(false);   // approval-sheet export overlay
  const toggleSel = (id) => setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const clearSel = () => setSelected(new Set());

  const rooms = useMemo(
    () => Array.from(new Set([...(areas || []), ...items.map((i) => i.room).filter(Boolean)])),
    [areas, items]
  );
  const types = useMemo(
    () => Array.from(new Set([...BASE_TYPES, ...items.map((i) => i.itemType).filter(Boolean)])),
    [items]
  );

  /* ---------- filtering ---------- */
  const visible = useMemo(() => {
    let v = items;
    if (clientView) v = v.filter((i) => CLIENT_STATUSES.includes(i.status));
    if (fRoom !== "All") v = v.filter((i) => i.room === fRoom);
    if (fStatus !== "All") v = v.filter((i) => i.status === fStatus);
    if (search.trim()) {
      const q = search.toLowerCase();
      v = v.filter((i) =>
        [i.name, i.retailer, i.finish, i.room, i.itemType, i.sku].join(" ").toLowerCase().includes(q)
      );
    }
    return v;
  }, [items, clientView, fRoom, fStatus, search]);

  const grouped = useMemo(() => {
    const g = {};
    // when simply browsing, show managed areas as groups even if empty, in their saved order
    const pureBrowse = !search.trim() && fStatus === "All" && !clientView;
    if (pureBrowse) (fRoom === "All" ? rooms : [fRoom]).forEach((r) => { if (r) g[r] = []; });
    visible.forEach((i) => { (g[i.room || "Unassigned"] ||= []).push(i); });
    return g;
  }, [visible, rooms, fRoom, fStatus, search, clientView]);

  const totals = useMemo(() => {
    const useNet = !clientView;
    const sub = visible.reduce((s, i) => s + lineTotal(i, retailerDiscounts, useNet), 0);
    const retail = visible.reduce((s, i) => s + lineTotal(i, retailerDiscounts, false), 0);
    const tax = sub * TAX_RATE;
    return { sub, tax, total: sub + tax, count: visible.length, savings: retail - sub };
  }, [visible, retailerDiscounts, clientView]);

  /* ---------- notifications ---------- */
  const addNotif = (audience, type, text, itemId) =>
    onNotifs((prev) => [{ id: uid(), audience, type, text, itemId: itemId || null, date: nowTs(), read: false }, ...prev].slice(0, 200));
  const markRead = (audience) => onNotifs((prev) => prev.map((n) => (n.audience === audience ? { ...n, read: true } : n)));

  /* ---------- mutations ---------- */
  const update = (id, patch) => {
    if (patch.status === "Proposed to client") {
      const old = items.find((i) => i.id === id);
      if (old && old.status !== "Proposed to client") addNotif("client", "proposed", `“${old.name || "An item"}” is ready for your review.`, id);
    }
    setItems((p) => p.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  };
  const remove = (id) => setItems((p) => p.filter((i) => i.id !== id));

  // item comments + client approval decisions (with cross-party notifications)
  const addComment = (id, by, text) => {
    const t = (text || "").trim(); if (!t) return;
    setItems((p) => p.map((i) => (i.id === id ? { ...i, comments: [...(i.comments || []), { id: uid(), by, text: t, date: nowTs() }] } : i)));
    const it = items.find((i) => i.id === id); const label = it?.name || "an item";
    if (by === "client") addNotif("designer", "comment", `Client commented on “${label}”: ${t.slice(0, 90)}`, id);
    else addNotif("client", "comment", `Your designer replied on “${label}”.`, id);
  };
  const clientDecide = (id, decision) => {
    setItems((p) => p.map((i) => (i.id === id
      ? { ...i, approval: { state: decision === "approve" ? "Approved" : "Rejected", date: today() }, status: decision === "approve" ? "Client approved" : i.status }
      : i)));
    const it = items.find((i) => i.id === id); const label = it?.name || "an item";
    addNotif("designer", "approval", decision === "approve" ? `Client approved “${label}”.` : `Client requested changes to “${label}”.`, id);
  };

  /* ---------- areas (parts of the house) ---------- */
  const addArea = (name) => {
    const n = (name || "").trim();
    if (!n) return;
    onAreas((prev) => (prev.includes(n) ? prev : [...prev, n]));
  };
  const renameArea = (oldName, newName) => {
    const n = (newName || "").trim();
    if (!n || n === oldName) return;
    onAreas((prev) => prev.map((a) => (a === oldName ? n : a)));
    setItems((p) => p.map((i) => (i.room === oldName ? { ...i, room: n } : i)));
  };
  const deleteArea = (name) => {
    const count = items.filter((i) => i.room === name).length;
    if (count > 0 && !confirm(`"${name}" has ${count} item${count === 1 ? "" : "s"}. Delete this area and move them to "Unassigned"?`)) return;
    onAreas((prev) => prev.filter((a) => a !== name));
    if (count > 0) setItems((p) => p.map((i) => (i.room === name ? { ...i, room: "" } : i)));
    if (fRoom === name) setFRoom("All");
  };

  /* ---------- pinboard ---------- */
  const addPinImages = async (fileList) => {
    const reads = await readAttachments(fileList);
    const imgs = reads.filter((r) => !r.type || r.type.startsWith("image/"));
    onPinboard((prev) => [...imgs.map((r) => ({ id: r.id, src: r.dataUrl, name: r.name, note: "", tooBig: r.tooBig })), ...prev]);
  };
  const addPinUrl = (url) => {
    const u = (url || "").trim();
    if (!u) return;
    onPinboard((prev) => [{ id: uid(), src: u, name: "", note: "" }, ...prev]);
  };
  const updatePin = (id, patch) => onPinboard((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  const removePin = (id) => onPinboard((prev) => prev.filter((p) => p.id !== id));

  /* ---------- files / shared ---------- */
  const addFiles = async (fileList, by) => {
    const reads = await readAttachments(fileList, { sharedWithClient: by === "client", uploadedBy: by });
    onFiles((prev) => [...reads, ...prev]);
    if (by === "client" && reads.length) addNotif("designer", "file", reads.length === 1 ? `Client uploaded a file: ${reads[0].name}.` : `Client uploaded ${reads.length} files.`);
  };
  const toggleShared = (id) => {
    const f = files.find((x) => x.id === id);
    if (f && !f.sharedWithClient) addNotif("client", "file", `A file was shared with you: ${f.name}.`);
    onFiles((prev) => prev.map((x) => (x.id === id ? { ...x, sharedWithClient: !x.sharedWithClient } : x)));
  };
  const removeFile = (id) => onFiles((prev) => prev.filter((f) => f.id !== id));

  /* ---------- presentations ---------- */
  const newPage = (title = "New page") => ({ id: uid(), title, notes: "", images: [], productIds: [], pins: [] });
  const createPres = (title) => {
    const p = { id: uid(), title: title || "Untitled presentation", createdAt: today(), published: false, pages: [newPage("Cover")] };
    onPres((prev) => [p, ...prev]);
    return p.id;
  };
  const updatePres = (id, patch) => {
    if (patch.published) { const old = presentations.find((p) => p.id === id); if (old && !old.published) addNotif("client", "presentation", `New presentation shared: “${old.title}”.`); }
    onPres((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  };
  const deletePres = (id) => onPres((prev) => prev.filter((p) => p.id !== id));
  const mutatePages = (presId, fn) => onPres((prev) => prev.map((p) => (p.id === presId ? { ...p, pages: fn(p.pages) } : p)));
  const addPage = (presId) => mutatePages(presId, (pg) => [...pg, newPage(`Page ${pg.length + 1}`)]);
  const updatePage = (presId, pageId, patch) => mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId ? { ...p, ...patch } : p)));
  const removePage = (presId, pageId) => mutatePages(presId, (pg) => pg.filter((p) => p.id !== pageId));
  const movePage = (presId, idx, dir) => mutatePages(presId, (pg) => {
    const j = idx + dir; if (j < 0 || j >= pg.length) return pg;
    const n = [...pg]; [n[idx], n[j]] = [n[j], n[idx]]; return n;
  });
  const addPageImages = async (presId, pageId, fileList) => {
    const reads = await readAttachments(fileList);
    const imgs = reads.filter((r) => !r.type || r.type.startsWith("image/")).map((r) => ({ id: r.id, src: r.dataUrl, tooBig: r.tooBig }));
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId ? { ...p, images: [...p.images, ...imgs] } : p)));
  };
  const addPageImageSrc = (presId, pageId, src) => {
    if (!src.trim()) return;
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId ? { ...p, images: [...p.images, { id: uid(), src: src.trim() }] } : p)));
  };
  const removePageImage = (presId, pageId, imgId) =>
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId ? { ...p, images: p.images.filter((i) => i.id !== imgId), pins: (p.pins || []).filter((pn) => pn.imgId !== imgId) } : p)));
  const togglePageProduct = (presId, pageId, itemId) =>
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId
      ? { ...p, productIds: p.productIds.includes(itemId) ? p.productIds.filter((x) => x !== itemId) : [...p.productIds, itemId],
          pins: p.productIds.includes(itemId) ? (p.pins || []).filter((pn) => pn.itemId !== itemId) : (p.pins || []) }
      : p)));
  const addPin = (presId, pageId, imgId, itemId, x, y) =>
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId
      ? { ...p, pins: [...(p.pins || []), { id: uid(), imgId, itemId, x, y }], productIds: p.productIds.includes(itemId) ? p.productIds : [...p.productIds, itemId] }
      : p)));
  const removeProdPin = (presId, pageId, pinId) =>
    mutatePages(presId, (pg) => pg.map((p) => (p.id === pageId ? { ...p, pins: (p.pins || []).filter((pn) => pn.id !== pinId) } : p)));

  const startAddManual = () => setDraft(mk({ status: "Idea" }));
  const startAddByUrl = async () => {
    const url = urlInput.trim();
    if (!url) return;
    const d = mk({ url, retailer: retailerFromUrl(url), status: "Idea" });
    setDraft(d); setUrlInput(""); setFetching(true);
    const data = await fetchProductData(url);
    setFetching(false);
    setDraft((cur) => {
      if (!cur || cur.id !== d.id) return cur;
      if (!data) return { ...cur, _fetchFailed: true };
      return {
        ...cur,
        name: data.name || cur.name,
        sku: data.sku || cur.sku,
        priceAuto: data.price != null && !isNaN(data.price) ? Number(data.price) : cur.priceAuto,
        dimensions: data.dimensions || cur.dimensions,
        availability: data.availability || cur.availability,
        imageUrl: data.imageUrl || cur.imageUrl,
        retailer: data.retailer || cur.retailer,
        lastChecked: today(),
        _fetched: true,
      };
    });
  };

  const commitDraft = () => {
    if (!draft) return;
    const { _fetched, _fetchFailed, ...clean } = draft;
    setItems((p) => [...p, clean]);
    setDraft(null);
  };

  const refreshItem = async (it) => {
    if (!it.url) return;
    setRefreshing(it.id);
    const data = await fetchProductData(it.url);
    setRefreshing(null);
    if (!data) { update(it.id, { lastChecked: today() }); return; }
    const newPrice = data.price != null && !isNaN(data.price) ? Number(data.price) : it.priceAuto;
    const changed = Number(newPrice) !== Number(it.priceAuto);
    update(it.id, {
      priceAuto: newPrice,
      availability: data.availability || it.availability,
      dimensions: data.dimensions || it.dimensions,
      imageUrl: data.imageUrl || it.imageUrl,
      lastChecked: today(),
      lastPriceChange: changed ? { old: Number(it.priceAuto), new: Number(newPrice), date: today() } : it.lastPriceChange,
    });
  };

  const resetData = () => { if (confirm("Clear this client's schedule and reload the sample data?")) { setItems(seed()); setRetailerDiscounts(seedDiscounts()); } };

  // attach receipt/invoice files to an item (stored as data URLs)
  const addReceipts = async (id, fileList) => {
    const files = Array.from(fileList || []);
    const reads = await Promise.all(files.map((f) => new Promise((res) => {
      if (f.size > 4 * 1024 * 1024) { res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: "", tooBig: true, addedAt: today() }); return; }
      const r = new FileReader();
      r.onload = () => res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: r.result, addedAt: today() });
      r.onerror = () => res(null);
      r.readAsDataURL(f);
    })));
    setItems((prev) => prev.map((i) => i.id === id ? { ...i, receipts: [...(i.receipts || []), ...reads.filter(Boolean)] } : i));
  };
  const removeReceipt = (id, rid) =>
    setItems((prev) => prev.map((i) => i.id === id ? { ...i, receipts: (i.receipts || []).filter((r) => r.id !== rid) } : i));

  // scan Gmail for order/shipping info on one item
  const scanInbox = async (it) => {
    setScanning(it.id);
    const info = await fetchOrderInfo(it, client);
    setScanning(null);
    if (!info || !info.found) { update(it.id, { lastScanned: today(), scanResult: "none" }); return; }
    const patch = { lastScanned: today(), scanResult: "found" };
    if (info.orderNumber) patch.orderNumber = info.orderNumber;
    if (info.carrier) patch.carrier = info.carrier;
    if (info.trackingNumber) patch.trackingNumber = info.trackingNumber;
    if (info.estimatedDelivery) patch.eta = info.estimatedDelivery;
    if (info.orderDate) patch.orderDate = info.orderDate;
    if (info.statusGuess && info.statusGuess !== "Unknown" && STATUSES.includes(info.statusGuess)) patch.status = info.statusGuess;
    update(it.id, patch);
  };
  const scanAllOrders = async () => {
    setScanAll(true);
    const targets = items.filter((i) => ["Ordered", "Shipped", "Proposed to client", "Client approved"].includes(i.status));
    for (const it of targets) { await scanInbox(it); }
    setScanAll(false);
  };

  // fetch live shipment status from tracking number / URL
  const trackShipment = async (it) => {
    setTracking(it.id);
    const info = await fetchShipmentStatus(it);
    setTracking(null);
    if (!info || !info.found) { update(it.id, { shipChecked: today(), shipStatus: it.shipStatus || "" }); return; }
    const patch = { shipChecked: today() };
    if (info.status) patch.shipStatus = info.status + (info.location ? ` · ${info.location}` : "");
    if (info.estimatedDelivery) patch.eta = info.estimatedDelivery;
    if (info.delivered) patch.status = "Delivered";
    else if (info.status && /transit|shipped|out for delivery/i.test(info.status) && it.status === "Ordered") patch.status = "Shipped";
    update(it.id, patch);
  };

  const SECTIONS = (clientMode
    ? [["overview", "Overview", Home], ["furnishings", "Selections", Package], ["pinboard", "Pinboard", LayoutGrid], ["presentations", "Presentations", Presentation], ["files", "Files", Folder]]
    : [["overview", "Overview", Home], ["furnishings", "Schedule", Package], ["contractors", "Contractors", Users], ["finance", "Finance", Wallet], ["pinboard", "Pinboard", LayoutGrid], ["presentations", "Presentations", Presentation], ["files", "Files", Folder]]);
  const sectionTitle = { overview: "Overview", furnishings: clientMode ? "Selections" : "The Schedule", contractors: "Contractors & Trades", finance: "Finance", pinboard: "Design Pinboard", presentations: "Presentations", files: clientMode ? "Files" : "Files & Sharing" }[section];

  return (
    <div className="app proj-shell">
      <Style />
      <aside className="proj-nav">
        {clientMode
          ? <div className="client-badge"><Lock size={12} /> Client view</div>
          : <button className="back-link" onClick={onBack}><ChevronLeft size={14} /> All projects</button>}
        <div className="proj-id">
          <div className="proj-avatar font-display">{(client?.clientName || client?.projectName || "?").charAt(0)}</div>
          <div className="proj-id-text">
            <div className="proj-name">{client?.projectName}</div>
            {client?.clientName && <div className="proj-client">{client.clientName}</div>}
          </div>
        </div>
        <nav className="nav">
          {SECTIONS.map(([key, label, Icon]) => (
            <button key={key} className={"nav-item" + (section === key ? " on" : "")} onClick={() => setSection(key)}>
              <Icon size={16} /> {label}
            </button>
          ))}
        </nav>
        <div className="nav-foot">
          {clientMode ? (
            <button className="nav-item" onClick={onSignOut}><LogOut size={16} /> Sign out</button>
          ) : (
            <>
              <button className={"nav-item" + (clientView ? " on" : "")} onClick={() => setClientView((v) => !v)}>
                {clientView ? <Eye size={16} /> : <EyeOff size={16} />} {clientView ? "Client preview on" : "Client preview"}
              </button>
              <button className="nav-item" onClick={() => setShowAccess(true)}><KeyRound size={16} /> Client access</button>
              <button className="nav-item" onClick={resetData}><RotateCcw size={16} /> Reset sample data</button>
            </>
          )}
        </div>
      </aside>

      {showAccess && !clientMode && (
        <ClientAccessPanel client={client} onUpdateClient={onUpdateClient} onPreview={onPreviewClientLogin} onClose={() => setShowAccess(false)} />
      )}

      <div className="proj-content">
      {/* ---------------- header ---------------- */}
      <header className="hdr">
        <div className="hdr-row">
          <div>
            <div className="eyebrow">{client?.projectName || "Project"}{client?.clientName ? ` · ${client.clientName}` : ""}</div>
            <h1 className="font-display title">{"{ " + sectionTitle + " }"}</h1>
          </div>
          <div className="hdr-actions">
            {section === "furnishings" && (
              <div className="seg">
                <button className={view === "list" ? "on" : ""} onClick={() => setView("list")}><List size={15} /> List</button>
                <button className={view === "cards" ? "on" : ""} onClick={() => setView("cards")}><LayoutGrid size={15} /> Cards</button>
              </div>
            )}
            <NotificationsBell notifications={notifications} audience={clientMode ? "client" : "designer"} onMarkRead={markRead} onGo={(n) => { if (n.itemId) { setSection("furnishings"); setView("cards"); } else if (n.type === "presentation") { setSection("presentations"); } else if (n.type === "file") { setSection("files"); } }} />
          </div>
        </div>

        {/* totals */}
        {section === "furnishings" && (
          <div className="totals">
            <Totals label={fRoom === "All" && fStatus === "All" && !search ? "Project total" : "Filtered total"}
                    sub={totals.sub} tax={totals.tax} total={totals.total} count={totals.count}
                    savings={totals.savings} clientView={clientView} />
          </div>
        )}

        {clientView && section !== "overview" && (
          <div className="banner">
            <Eye size={14} /> Client preview — this is what your client sees. Internal cost &amp; notes are hidden.
          </div>
        )}
      </header>

      {section === "overview" && (
        <ProjectOverview client={client} items={items} discounts={discounts} contractors={contractors}
                         pinboard={pinboard} files={files} presentations={presentations} onNavigate={setSection}
                         clientMode={clientMode} onCover={onUpdateClient ? (v) => onUpdateClient({ coverPhoto: v }) : null} />
      )}

      {section === "presentations" && (
        <Presentations clientView={clientView} presentations={presentations} items={items} discounts={discounts} pinboard={pinboard}
                       client={client} onPreview={setPreview}
                       onCreate={createPres} onUpdate={updatePres} onDelete={deletePres}
                       onAddPage={addPage} onUpdatePage={updatePage} onRemovePage={removePage} onMovePage={movePage}
                       onAddImages={addPageImages} onAddImageSrc={addPageImageSrc} onRemoveImage={removePageImage} onToggleProduct={togglePageProduct}
                       onAddPin={addPin} onRemovePin={removeProdPin} />
      )}

      {section === "pinboard" && (
        <Pinboard clientView={clientView} pinboard={pinboard} onAddImages={addPinImages} onAddUrl={addPinUrl} onUpdate={updatePin} onRemove={removePin} onPreview={setPreview} />
      )}

      {section === "files" && (
        <FilesShared clientView={clientView} readOnly={portalReadOnly} files={files} onAdd={addFiles} onToggleShared={toggleShared} onRemove={removeFile} />
      )}


      {/* ---------------- toolbar ---------------- */}
      {section === "furnishings" && !clientView && (
        <div className="toolbar">
          <div className="url-add">
            <Link2 size={16} className="url-ico" />
            <input className="input url-input" placeholder="Paste a product link to auto-fill…"
                   value={urlInput} onChange={(e) => setUrlInput(e.target.value)}
                   onKeyDown={(e) => e.key === "Enter" && startAddByUrl()} />
            <button className="btn btn-primary" onClick={startAddByUrl} disabled={!urlInput.trim()}>
              {fetching ? <Loader2 size={15} className="spin" /> : <Plus size={15} />} Auto-fill
            </button>
            <button className="btn btn-ghost" onClick={startAddManual}><Plus size={15} /> Manual</button>
            <button className={"btn " + (showDiscounts ? "btn-primary" : "btn-ghost")} onClick={() => { setShowDiscounts((v) => !v); setShowAreas(false); }}>
              <Percent size={15} /> Discounts
            </button>
            <button className={"btn " + (showAreas ? "btn-primary" : "btn-ghost")} onClick={() => { setShowAreas((v) => !v); setShowDiscounts(false); }}>
              <Building2 size={15} /> Areas
            </button>
            <button className="btn btn-ghost" onClick={scanAllOrders} disabled={scanAll} title="Search Gmail for order & shipping updates">
              {scanAll ? <Loader2 size={15} className="spin" /> : <Inbox size={15} />} Scan inbox
            </button>
          </div>
          {showDiscounts && (
            <DiscountPanel items={items} discounts={retailerDiscounts}
                           onChange={(retailer, val) => setRetailerDiscounts((d) => ({ ...d, [retailer]: val }))}
                           onRemove={(retailer) => setRetailerDiscounts((d) => { const n = { ...d }; delete n[retailer]; return n; })}
                           onClose={() => setShowDiscounts(false)} />
          )}
          {showAreas && (
            <AreaPanel rooms={rooms} items={items} onAdd={addArea} onRename={renameArea} onDelete={deleteArea} onClose={() => setShowAreas(false)} />
          )}
        </div>
      )}

      {/* filters */}
      {section === "furnishings" && (
      <div className="filters">
        <div className="search-wrap">
          <Search size={15} className="search-ico" />
          <input className="input" placeholder="Search items, retailers, finishes…" value={search}
                 onChange={(e) => setSearch(e.target.value)} />
        </div>
        <select className="select" value={fRoom} onChange={(e) => setFRoom(e.target.value)}>
          <option>All</option>{rooms.map((r) => <option key={r}>{r}</option>)}
        </select>
        <select className="select" value={fStatus} onChange={(e) => setFStatus(e.target.value)}>
          <option>All</option>{STATUSES.map((s) => <option key={s}>{s}</option>)}
        </select>
      </div>
      )}

      {/* ---------------- draft / review panel ---------------- */}
      {section === "furnishings" && draft && (
        <DraftPanel draft={draft} setDraft={setDraft} fetching={fetching}
                    rooms={rooms} types={types} onCommit={commitDraft} onCancel={() => setDraft(null)} onPreview={setPreview} />
      )}

      {/* ---------------- schedule ---------------- */}
      {section === "furnishings" && (
      <main className="schedule">
        {Object.keys(grouped).length === 0 && (
          <div className="empty"><Package size={28} /><p>No items match. Adjust filters or paste a product link.</p></div>
        )}

        {Object.keys(grouped).length > 0 && view === "list" && (
          <ScheduleTable grouped={grouped} clientView={clientView} refreshing={refreshing} rd={retailerDiscounts}
                         scanning={scanning} tracking={tracking} selectable={!clientView} selected={selected} onToggleSel={toggleSel}
                         onScan={scanInbox} onTrack={trackShipment} onPreview={setPreview}
                         onAddReceipts={addReceipts} onRemoveReceipt={removeReceipt}
                         onUpdate={update} onRemove={remove} onRefresh={refreshItem} />
        )}

        {Object.keys(grouped).length > 0 && view === "cards" && Object.entries(grouped).map(([room, list]) => {
          const sub = list.reduce((s, i) => s + lineTotal(i, retailerDiscounts, !clientView), 0);
          const isCol = collapsed[room];
          return (
            <section key={room} className="room">
              <button className="room-head" onClick={() => setCollapsed((c) => ({ ...c, [room]: !c[room] }))}>
                {isCol ? <ChevronRight size={18} /> : <ChevronDown size={18} />}
                <span className="font-display room-name">{room}</span>
                <span className="room-count">{list.length} {list.length === 1 ? "item" : "items"}</span>
                <span className="room-sub">{money(sub)} <em>+ {money(sub * TAX_RATE)} tax</em></span>
              </button>
              {!isCol && (
                <div className="cards">
                  {list.length === 0
                    ? <div className="area-empty">No items in this area yet.</div>
                    : list.map((it) => (
                    <ItemCard key={it.id} it={it} clientView={clientView} rd={retailerDiscounts}
                              selectable={!clientView} sel={selected.has(it.id)} onSel={() => toggleSel(it.id)}
                              refreshing={refreshing === it.id} scanning={scanning === it.id} tracking={tracking === it.id}
                              onScan={() => scanInbox(it)} onTrack={() => trackShipment(it)}
                              onAddReceipts={(fl) => addReceipts(it.id, fl)} onRemoveReceipt={(rid) => removeReceipt(it.id, rid)}
                              onComment={(by, text) => addComment(it.id, by, text)} onDecide={(d) => clientDecide(it.id, d)}
                              onPreview={setPreview} readOnly={portalReadOnly}
                              onUpdate={(p) => update(it.id, p)} onRemove={() => remove(it.id)}
                              onRefresh={() => refreshItem(it)} />
                  ))}
                </div>
              )}
            </section>
          );
        })}
      </main>
      )}

      {/* ---------------- contractors ---------------- */}
      {section === "contractors" && (
        <Contractors clientView={clientView} contractors={contractors} onContractors={onContractors} />
      )}

      {section === "finance" && !clientMode && (
        <Finance payments={payments} onPayments={onPayments} items={items} discounts={retailerDiscounts} contractors={contractors} />
      )}

      {/* selection bar for export */}
      {section === "furnishings" && !clientView && selected.size > 0 && (
        <div className="sel-bar">
          <span className="sel-count">{selected.size} selected</span>
          <button className="btn btn-ghost" onClick={() => setSelected(new Set(items.map((i) => i.id)))}>Select all</button>
          <button className="btn btn-ghost" onClick={clearSel}>Clear</button>
          <div style={{ flex: 1 }} />
          <button className="btn btn-ghost" onClick={() => exportScheduleXlsx(items.filter((i) => selected.has(i.id)), client, retailerDiscounts)}><Download size={15} /> Export Excel</button>
          <button className="btn btn-primary" onClick={() => setExporting(true)}><FileText size={15} /> Export approval sheet</button>
        </div>
      )}

      {exporting && (
        <ExportOverlay title="Selections for approval" filename={`${(client?.projectName || "selections").replace(/\s+/g, "_")}_approval`} onClose={() => setExporting(false)}>
          <ApprovalSheet client={client} items={items.filter((i) => selected.has(i.id))} discounts={retailerDiscounts} />
        </ExportOverlay>
      )}

      {preview && <Lightbox src={preview} onClose={() => setPreview(null)} />}

      <footer className="foot">
        {client?.projectName} · saved to this workspace
      </footer>
      </div>
    </div>
  );
}

/* ----------------------------- Image preview (lightbox) ----------------------------- */
function Lightbox({ src, onClose }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  return (
    <div className="lightbox" onClick={onClose}>
      <button className="lightbox-x" onClick={onClose} aria-label="Close" title="Close"><X size={20} /></button>
      <img className="lightbox-img" src={src} alt="" onClick={(e) => e.stopPropagation()} />
    </div>
  );
}

/* ----------------------------- Notifications ----------------------------- */
function NotificationsBell({ notifications, audience, onMarkRead, onGo }) {
  const [open, setOpen] = useState(false);
  const mine = notifications.filter((n) => n.audience === audience);
  const unread = mine.filter((n) => !n.read).length;
  const toggle = () => { const willOpen = !open; setOpen(willOpen); if (willOpen && unread) onMarkRead(audience); };
  const ico = (t) => t === "approval" ? <Check size={14} /> : t === "comment" ? <MessageSquare size={14} /> : t === "file" ? <Folder size={14} /> : t === "presentation" ? <Presentation size={14} /> : t === "proposed" ? <Package size={14} /> : <Bell size={14} />;
  return (
    <div className="notif-wrap">
      <button className="notif-bell" onClick={toggle} title="Notifications"><Bell size={18} />{unread > 0 && <span className="notif-dot">{unread > 9 ? "9+" : unread}</span>}</button>
      {open && (
        <>
          <div className="notif-overlay" onClick={() => setOpen(false)} />
          <div className="notif-panel">
            <div className="notif-hd font-display">Notifications</div>
            <div className="notif-list">
              {mine.length === 0 && <div className="notif-empty">You’re all caught up.</div>}
              {mine.slice(0, 40).map((n) => (
                <button key={n.id} className={"notif-item" + (n.read ? "" : " unread")} onClick={() => { if (onGo) onGo(n); setOpen(false); }}>
                  <span className="notif-ico">{ico(n.type)}</span>
                  <span className="notif-body"><span className="notif-text">{n.text}</span><span className="notif-when">{fmtWhen(n.date)}</span></span>
                </button>
              ))}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

/* ----------------------------- Export overlay (PDF via print + Word) ----------------------------- */
function ExportOverlay({ title, filename, onClose, children }) {
  const ref = useRef(null);
  const doDoc = () => {
    const html = ref.current ? ref.current.innerHTML : "";
    const doc = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><style>body{font-family:Georgia,'Times New Roman',serif;color:#222;margin:24px}img{max-width:100%;height:auto}.pg{page-break-after:always;margin-bottom:28px}h1,h2,h3{font-family:Georgia,serif}a{color:#1155cc}table{border-collapse:collapse}</style></head><body>${html}</body></html>`;
    const blob = new Blob(["\ufeff", doc], { type: "application/msword" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = (filename || "export") + ".doc";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  };
  return (
    <div className="export-modal" role="dialog" aria-modal="true">
      <div className="export-bar">
        <div className="export-title font-display">{title}</div>
        <div className="export-actions">
          <button className="btn" onClick={doDoc} title="Download an editable Word document"><FileText size={15} /> Word (.doc)</button>
          <button className="btn btn-primary" onClick={() => window.print()} title="Use your browser's “Save as PDF”"><Download size={15} /> Print / Save PDF</button>
          <button className="icon-btn" onClick={onClose} title="Close"><X size={17} /></button>
        </div>
      </div>
      <div className="export-scroll"><div className="print-area" ref={ref}>{children}</div></div>
    </div>
  );
}

/* ----------------------------- Approval sheet ----------------------------- */
function ApprovalSheet({ client, items, discounts }) {
  const byRoom = {};
  items.forEach((i) => { (byRoom[i.room || "Other"] ||= []).push(i); });
  return (
    <div className="doc">
      <div className="doc-hd">
        <div>
          <div className="doc-eyebrow">Selections for approval</div>
          <h1 className="font-display">{client?.projectName || "Project"}</h1>
          {client?.clientName && <div className="doc-sub">Prepared for {client.clientName}</div>}
        </div>
        <div className="doc-date">{new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}</div>
      </div>
      {items.length === 0 && <p className="muted">No items selected.</p>}
      {Object.entries(byRoom).map(([room, list]) => (
        <div key={room} className="doc-room">
          <h2 className="font-display">{room}</h2>
          {list.map((it) => {
            const net = netEach(it, discounts);
            return (
              <div key={it.id} className="doc-item">
                <div className="doc-thumb">{it.imageUrl ? <img src={it.imageUrl} alt="" /> : <div className="thumb-ph"><ImageIcon size={20} /></div>}</div>
                <div className="doc-meta">
                  <div className="doc-name">{it.name || "Item"}</div>
                  {[it.retailer, it.itemType].filter(Boolean).length > 0 && <div className="doc-line">{[it.retailer, it.itemType].filter(Boolean).join(" · ")}</div>}
                  {it.finish && <div className="doc-line">Finish / selection: {it.finish}</div>}
                  {it.dimensions && <div className="doc-line">Dimensions: {it.dimensions}</div>}
                  <div className="doc-line">Qty {it.qty} · {money(net)} each · {money(net * (Number(it.qty) || 1))} total</div>
                  {it.url && <a className="doc-link" href={it.url}>{it.url}</a>}
                </div>
                <div className="doc-signoff">
                  <div className="sign-opt">☐ Approve</div>
                  <div className="sign-opt">☐ Decline</div>
                  <div className="sign-opt">☐ Revise</div>
                </div>
              </div>
            );
          })}
        </div>
      ))}
      <div className="doc-footer">Approved by ______________________________   Date __________________</div>
    </div>
  );
}

/* ----------------------------- Presentations ----------------------------- */
function PresentationDoc({ pres, items, discounts, client, clientView }) {
  const [openPin, setOpenPin] = useState(null);
  return (
    <div className="pres-doc" onClick={() => setOpenPin(null)}>
      {pres.pages.map((pg, idx) => {
        const pins = pg.pins || [];
        const pinNo = {}; pins.forEach((pn, i) => { pinNo[pn.id] = i + 1; });
        const pinnedIds = pins.map((p) => p.itemId);
        const ordered = [...pins.map((p) => ({ itemId: p.itemId, no: pinNo[p.id] })), ...pg.productIds.filter((id) => !pinnedIds.includes(id)).map((id) => ({ itemId: id }))];
        return (
          <section key={pg.id} className="pg pres-slide">
            <div className="pres-slide-hd">
              <h2 className="font-display">{pg.title || `Page ${idx + 1}`}</h2>
              <span className="pres-slide-proj">{client?.projectName}</span>
            </div>
            {pg.images.length > 0 && (
              <div className={"pres-slide-imgs n" + Math.min(pg.images.length, 4)}>
                {pg.images.map((im) => {
                  const imgPins = pins.filter((p) => p.imgId === im.id);
                  return (
                    <div key={im.id} className="pres-stage">
                      {im.src ? <img src={im.src} alt="" /> : <div className="thumb-ph"><ImageIcon size={18} /></div>}
                      {imgPins.map((pn) => {
                        const it = items.find((i) => i.id === pn.itemId);
                        const open = openPin === pn.id;
                        return (
                          <div key={pn.id} className="pin-dot-wrap" style={{ left: pn.x + "%", top: pn.y + "%" }}>
                            <button className={"pin-dot" + (open ? " on" : "")} onClick={(e) => { e.stopPropagation(); setOpenPin(open ? null : pn.id); }} title={it ? it.name : ""}>{pinNo[pn.id]}</button>
                            {open && it && (
                              <div className={"pin-pop" + (pn.x > 60 ? " left" : "") + (pn.y > 55 ? " up" : "")} onClick={(e) => e.stopPropagation()}>
                                {it.imageUrl && <img className="pin-pop-img" src={it.imageUrl} alt="" />}
                                <div className="pin-pop-info">
                                  <strong>{it.name || "Item"}</strong>
                                  {[it.retailer, it.itemType].filter(Boolean).length > 0 && <span>{[it.retailer, it.itemType].filter(Boolean).join(" · ")}</span>}
                                  {it.finish && <span>Finish: {it.finish}</span>}
                                  {it.dimensions && <span>{it.dimensions}</span>}
                                  {!clientView && <span className="pin-pop-price">{money(netEach(it, discounts))} · qty {it.qty}</span>}
                                  {it.url && <a href={it.url} target="_blank" rel="noreferrer">View product →</a>}
                                </div>
                                <button className="pin-pop-x" onClick={() => setOpenPin(null)} title="Close"><X size={12} /></button>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            )}
            {ordered.length > 0 && (
              <div className="pres-slide-prods">
                {ordered.map(({ itemId, no }) => {
                  const it = items.find((i) => i.id === itemId);
                  if (!it) return null;
                  const net = netEach(it, discounts);
                  return (
                    <div key={itemId} className="pres-prod">
                      {no != null && <span className="pres-prod-no">{no}</span>}
                      {it.imageUrl ? <img src={it.imageUrl} alt="" /> : <div className="thumb-ph"><ImageIcon size={18} /></div>}
                      <div className="pres-prod-info">
                        <strong>{it.name || "Item"}</strong>
                        {[it.retailer, it.finish].filter(Boolean).length > 0 && <span>{[it.retailer, it.finish].filter(Boolean).join(" · ")}</span>}
                        {it.dimensions && <span>{it.dimensions}</span>}
                        {!clientView && <span>{money(net)} · qty {it.qty}</span>}
                        {it.url && <a href={it.url}>View product</a>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
            {pg.notes && <p className="pres-slide-note">{pg.notes}</p>}
          </section>
        );
      })}
    </div>
  );
}

function PresPageEditor({ page, idx, total, items, pinboard, onUpdate, onRemove, onMove, onAddImages, onAddImageSrc, onRemoveImage, onToggleProduct, onAddPin, onRemovePin }) {
  const [url, setUrl] = useState("");
  const [showPin, setShowPin] = useState(false);
  const [showProd, setShowProd] = useState(false);
  const [prodQuery, setProdQuery] = useState("");
  const [placing, setPlacing] = useState("");   // itemId being placed
  const [drag, setDrag] = useState(false);
  const pins = page.pins || [];
  const pinNo = {}; pins.forEach((pn, i) => { pinNo[pn.id] = i + 1; });
  const placingItem = items.find((i) => i.id === placing);

  useEffect(() => {
    if (!placing) return;
    const onKey = (e) => { if (e.key === "Escape") setPlacing(""); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [placing]);

  const dropHandlers = {
    onDragOver: (e) => { e.preventDefault(); setDrag(true); },
    onDragLeave: () => setDrag(false),
    onDrop: (e) => { e.preventDefault(); setDrag(false); const fs = Array.from(e.dataTransfer.files || []).filter((f) => f.type.startsWith("image/")); if (fs.length) onAddImages(fs); },
    onPaste: (e) => {
      const imgs = Array.from(e.clipboardData?.items || []).filter((it) => it.type.startsWith("image/")).map((it) => it.getAsFile()).filter(Boolean);
      if (imgs.length) { onAddImages(imgs); e.preventDefault(); return; }
      const t = e.clipboardData?.getData("text"); if (t && /^https?:\/\//.test(t.trim())) { onAddImageSrc(t.trim()); e.preventDefault(); }
    },
  };
  const placeOnImage = (imgId, e) => {
    if (!placing) return;
    const r = e.currentTarget.getBoundingClientRect();
    const x = Math.max(2, Math.min(98, ((e.clientX - r.left) / r.width) * 100));
    const y = Math.max(2, Math.min(98, ((e.clientY - r.top) / r.height) * 100));
    onAddPin(imgId, placing, Math.round(x), Math.round(y));
    setPlacing("");
  };

  return (
    <div className="pres-page">
      <div className="pres-page-hd">
        <span className="pres-page-num">Page {idx + 1}</span>
        <input className="input" value={page.title} onChange={(e) => onUpdate({ title: e.target.value })} placeholder="Page title — e.g. Living Room · Coffee Table Options" />
        <div className="pres-page-tools">
          <button className="icon-btn" disabled={idx === 0} onClick={() => onMove(-1)} title="Move up"><ArrowUp size={14} /></button>
          <button className="icon-btn" disabled={idx === total - 1} onClick={() => onMove(1)} title="Move down"><ArrowDown size={14} /></button>
          <button className="icon-btn" onClick={onRemove} title="Delete page"><Trash2 size={14} /></button>
        </div>
      </div>

      {placing && <div className="pin-place-bar"><Pin size={13} /> Click on an image below to pin <strong>{placingItem?.name || "item"}</strong>. <button className="link-mini" onClick={() => setPlacing("")}>Cancel</button></div>}

      <div className={"pres-drop" + (drag ? " drag" : "")} tabIndex={0} {...dropHandlers}>
        {page.images.length > 0 && (
          <div className="pres-edit-imgs">
            {page.images.map((im) => {
              const imgPins = pins.filter((p) => p.imgId === im.id);
              return (
                <div key={im.id} className={"pres-edit-stage" + (placing ? " placing" : "")} onClick={(e) => placeOnImage(im.id, e)}>
                  {im.src ? <img src={im.src} alt="" /> : <div className="pin-ph"><ImageIcon size={18} /></div>}
                  {imgPins.map((pn) => {
                    const it = items.find((i) => i.id === pn.itemId);
                    return (
                      <button key={pn.id} className="pin-dot edit" style={{ left: pn.x + "%", top: pn.y + "%" }}
                              onClick={(e) => { e.stopPropagation(); if (confirm(`Remove pin for “${it?.name || "item"}”?`)) onRemovePin(pn.id); }}
                              title={`${it?.name || "item"} — click to remove pin`}>{pinNo[pn.id]}</button>
                    );
                  })}
                  <button className="pin-x" onClick={(e) => { e.stopPropagation(); onRemoveImage(im.id); }} title="Remove image"><X size={12} /></button>
                </div>
              );
            })}
          </div>
        )}
        <div className="pres-add-row">
          <label className="pres-add-img"><Upload size={18} /><span>Upload</span><input type="file" multiple accept="image/*" style={{ display: "none" }} onChange={(e) => { onAddImages(e.target.files); e.target.value = ""; }} /></label>
          <span className="pres-drop-hint">Drag &amp; drop or paste an image here</span>
        </div>
      </div>

      <div className="pres-img-row">
        <input className="input" placeholder="…or paste an image URL" value={url} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { onAddImageSrc(url); setUrl(""); } }} />
        <button className="btn btn-ghost" onClick={() => { onAddImageSrc(url); setUrl(""); }} disabled={!url.trim()}>Add URL</button>
        {pinboard.length > 0 && <button className="btn btn-ghost" onClick={() => setShowPin((v) => !v)}><LayoutGrid size={14} /> From pinboard</button>}
      </div>
      {showPin && <div className="pres-pin-strip">{pinboard.map((p) => (<button key={p.id} className="pres-pin-thumb" onClick={() => onAddImageSrc(p.src)}>{p.src ? <img src={p.src} alt="" /> : <ImageIcon size={16} />}</button>))}</div>}

      {/* pin a product onto an image */}
      {page.images.length > 0 && items.length > 0 && (
        <div className="pres-pinsel">
          <Pin size={13} />
          <span>Pin a product onto an image:</span>
          <select className="select" value={placing} onChange={(e) => setPlacing(e.target.value)}>
            <option value="">Choose a product…</option>
            {items.map((it) => <option key={it.id} value={it.id}>{(it.name || "Item") + (it.room ? ` · ${it.room}` : "")}</option>)}
          </select>
        </div>
      )}

      <div className="pres-prod-hd"><span>Products on this page {pins.length > 0 && <em className="muted">· {pins.length} pinned</em>}{page.productIds.length > 0 && <em className="muted"> · {page.productIds.length} added</em>}</span><button className="btn btn-ghost" onClick={() => setShowProd((v) => !v)}>{showProd ? "Done" : "+ Add products"}</button></div>
      <div className="pres-prod-chips">
        {page.productIds.length === 0 && !showProd && <span className="muted" style={{ fontSize: 14.5 }}>No products yet — add them here or pin them on an image above.</span>}
        {page.productIds.map((pid) => { const it = items.find((i) => i.id === pid); if (!it) return null; const isPinned = pins.some((p) => p.itemId === pid); return <span key={pid} className="prod-chip">{isPinned && <Pin size={10} />}{it.name || "Item"}<button onClick={() => onToggleProduct(pid)}><X size={11} /></button></span>; })}
      </div>
      {showProd && (() => {
        const q = prodQuery.trim().toLowerCase();
        const filtered = items.filter((it) => !q || [it.name, it.room, it.itemType, it.retailer, it.sku].some((v) => (v || "").toLowerCase().includes(q)));
        const groups = {};
        filtered.forEach((it) => { (groups[it.room || "Unassigned"] ||= []).push(it); });
        const roomKeys = Object.keys(groups);
        return (
          <div className="pres-prod-panel">
            <div className="pres-prod-search">
              <Search size={15} />
              <input className="input" placeholder="Search by product, room, retailer, or SKU…" value={prodQuery} onChange={(e) => setProdQuery(e.target.value)} autoFocus />
              {prodQuery && <button className="icon-btn" onClick={() => setProdQuery("")} title="Clear search"><X size={14} /></button>}
            </div>
            <div className="pres-prod-list">
              {items.length === 0 && <div className="muted" style={{ fontSize: 14.5, padding: "10px 4px" }}>No items in the Schedule yet.</div>}
              {items.length > 0 && filtered.length === 0 && <div className="muted" style={{ fontSize: 14.5, padding: "10px 4px" }}>No products match “{prodQuery}”.</div>}
              {roomKeys.map((room) => (
                <div key={room} className="pres-prod-group">
                  <div className="pres-prod-room">{room} <em>{groups[room].length}</em></div>
                  {groups[room].map((it) => {
                    const on = page.productIds.includes(it.id);
                    const isPinned = pins.some((p) => p.itemId === it.id);
                    return (
                      <button key={it.id} className={"prod-pick" + (on ? " on" : "")} onClick={() => onToggleProduct(it.id)}>
                        {on ? <CheckSquare size={15} /> : <Square size={15} />}
                        <span className="prod-pick-name">{it.name || "Item"}</span>
                        {it.itemType && <em className="prod-pick-type">{it.itemType}</em>}
                        {isPinned && <Pin size={11} className="prod-pick-pin" />}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          </div>
        );
      })()}
      <textarea className="input pres-notes" rows={2} placeholder="Notes for this page (e.g. “Option 3 is my favorite — I love the black dipped legs”)" value={page.notes} onChange={(e) => onUpdate({ notes: e.target.value })} />
    </div>
  );
}

function Presentations({ clientView, presentations, items, discounts, pinboard, client, onPreview, onCreate, onUpdate, onDelete, onAddPage, onUpdatePage, onRemovePage, onMovePage, onAddImages, onAddImageSrc, onRemoveImage, onToggleProduct, onAddPin, onRemovePin }) {
  const [openId, setOpenId] = useState(null);
  const [exportId, setExportId] = useState(null);
  const open = presentations.find((p) => p.id === openId);
  const exportPres = presentations.find((p) => p.id === exportId);

  if (clientView) {
    const pub = presentations.filter((p) => p.published);
    return (
      <div className="pres-wrap">
        {pub.length === 0
          ? <div className="empty" style={{ padding: 60 }}><Presentation size={26} /><p>No presentations shared with you yet.</p></div>
          : <div className="pres-grid">{pub.map((p) => (
              <button key={p.id} className="pres-card" onClick={() => setExportId(p.id)}>
                <div className="pres-card-ico"><Presentation size={22} /></div>
                <div className="pres-card-title font-display">{p.title}</div>
                <div className="pres-card-meta">{p.pages.length} page{p.pages.length === 1 ? "" : "s"}</div>
              </button>))}</div>}
        {exportPres && <ExportOverlay title={exportPres.title} filename={exportPres.title.replace(/\s+/g, "_")} onClose={() => setExportId(null)}><PresentationDoc pres={exportPres} items={items} discounts={discounts} client={client} clientView /></ExportOverlay>}
      </div>
    );
  }

  if (!open) {
    return (
      <div className="pres-wrap">
        <div className="toolbar"><div className="url-add" style={{ justifyContent: "space-between" }}>
          <span style={{ flex: 1, fontSize: 15.5, color: "var(--ink-soft)" }}>Build option decks &amp; client presentations — pages with images, notes, and live products from the Schedule. Export to PDF or Word.</span>
          <button className="btn btn-primary" onClick={() => { const id = onCreate("Untitled presentation"); setOpenId(id); }}><Plus size={15} /> New presentation</button>
        </div></div>
        {presentations.length === 0
          ? <div className="empty" style={{ padding: 60 }}><Presentation size={26} /><p>No presentations yet. Create one to start your client deck.</p></div>
          : <div className="pres-grid">{presentations.map((p) => (
              <div key={p.id} className="pres-card">
                <button className="pres-card-ico" onClick={() => setOpenId(p.id)}><Presentation size={22} /></button>
                <div className="pres-card-title font-display" style={{ cursor: "pointer" }} onClick={() => setOpenId(p.id)}>{p.title}</div>
                <div className="pres-card-meta">{p.pages.length} page{p.pages.length === 1 ? "" : "s"} · {fmtDate(p.createdAt)}{p.published ? " · shared" : ""}</div>
                <div className="pres-card-actions">
                  <button className="btn btn-ghost" onClick={() => setOpenId(p.id)}>Open</button>
                  <button className="btn btn-ghost" onClick={() => setExportId(p.id)} title="Preview / export"><Download size={14} /></button>
                  <button className="icon-btn" onClick={() => { if (confirm("Delete this presentation?")) onDelete(p.id); }} title="Delete presentation"><Trash2 size={14} /></button>
                </div>
              </div>))}</div>}
        {exportPres && <ExportOverlay title={exportPres.title} filename={exportPres.title.replace(/\s+/g, "_")} onClose={() => setExportId(null)}><PresentationDoc pres={exportPres} items={items} discounts={discounts} client={client} /></ExportOverlay>}
      </div>
    );
  }

  return (
    <div className="pres-wrap">
      <div className="pres-editbar">
        <button className="btn btn-ghost" onClick={() => setOpenId(null)}><ChevronLeft size={15} /> All presentations</button>
        <input className="input pres-title-input" value={open.title} onChange={(e) => onUpdate(open.id, { title: e.target.value })} />
        <label className={"chip" + (open.published ? " on" : "")} style={{ cursor: "pointer" }}>
          <input type="checkbox" checked={!!open.published} onChange={(e) => onUpdate(open.id, { published: e.target.checked })} style={{ marginRight: 6 }} />
          Share with client
        </label>
        <button className="btn" onClick={() => onAddPage(open.id)}><Plus size={15} /> Add page</button>
        <button className="btn btn-primary" onClick={() => setExportId(open.id)}><Download size={15} /> Preview / Export</button>
      </div>
      <div className="pres-pages">
        {open.pages.map((pg, idx) => (
          <PresPageEditor key={pg.id} page={pg} idx={idx} total={open.pages.length} items={items} pinboard={pinboard}
            onUpdate={(patch) => onUpdatePage(open.id, pg.id, patch)} onRemove={() => onRemovePage(open.id, pg.id)} onMove={(dir) => onMovePage(open.id, idx, dir)}
            onAddImages={(fl) => onAddImages(open.id, pg.id, fl)} onAddImageSrc={(src) => onAddImageSrc(open.id, pg.id, src)}
            onRemoveImage={(iid) => onRemoveImage(open.id, pg.id, iid)} onToggleProduct={(iid) => onToggleProduct(open.id, pg.id, iid)}
            onAddPin={(imgId, itemId, x, y) => onAddPin(open.id, pg.id, imgId, itemId, x, y)} onRemovePin={(pinId) => onRemovePin(open.id, pg.id, pinId)} />
        ))}
      </div>
      {exportPres && <ExportOverlay title={open.title} filename={open.title.replace(/\s+/g, "_")} onClose={() => setExportId(null)}><PresentationDoc pres={open} items={items} discounts={discounts} client={client} /></ExportOverlay>}
    </div>
  );
}

/* ----------------------------- Project Overview ----------------------------- */
function ProjectOverview({ client, items, discounts, contractors, pinboard, files, presentations = [], onNavigate, clientMode = false, onCover }) {
  const net = items.reduce((s, i) => s + lineTotal(i, discounts, true), 0);
  const total = net * (1 + TAX_RATE);
  const awaiting = items.filter((i) => i.status === "Proposed to client" && i.approval.state === "Pending").length;
  const inTransit = items.filter((i) => ["Ordered", "Shipped"].includes(i.status)).length;
  const delivered = items.filter((i) => ["Delivered", "Installed"].includes(i.status)).length;
  const owed = contractors.reduce((s, c) => s + Math.max(0, balanceOf(c)), 0);
  const recent = pinboard.slice(0, 6);
  const sharedCount = files.filter((f) => f.sharedWithClient).length;

  const Stat = ({ label, value, sub, onClick }) => (
    <button className="ov-stat" onClick={onClick} style={{ cursor: onClick ? "pointer" : "default" }}>
      <div className="ov-stat-val font-display">{value}</div>
      <div className="ov-stat-label">{label}</div>
      {sub && <div className="ov-stat-sub">{sub}</div>}
    </button>
  );

  return (
    <div className="ov">
      {client?.coverPhoto && <div className="ov-cover"><img src={client.coverPhoto} alt="" /><div className="ov-cover-name font-display">{client.projectName}</div></div>}
      {!clientMode && onCover && (
        <div className="ov-cover-edit">
          <span className="ov-cover-label">{client?.coverPhoto ? "Cover photo" : "Add a cover photo of the house"}</span>
          <ImageInput value={client?.coverPhoto || ""} onChange={(v) => onCover(v)} />
        </div>
      )}
      <div className="ov-grid">
        <Stat label="Estimated total" value={money(total)} sub={`${items.length} items · incl. ~${Math.round(TAX_RATE * 100)}% tax`} onClick={() => onNavigate("furnishings")} />
        <Stat label="Awaiting client approval" value={awaiting} sub={awaiting ? "needs sign-off" : "all clear"} onClick={() => onNavigate("furnishings")} />
        <Stat label="In transit" value={inTransit} sub={`${delivered} delivered`} onClick={() => onNavigate("furnishings")} />
        <Stat label="Owed to contractors" value={money(owed)} sub={`${contractors.length} trades`} onClick={() => onNavigate("contractors")} />
      </div>

      <div className="ov-cols">
        <div className="ov-card">
          <div className="ov-card-head"><span className="font-display">Design pinboard</span><button className="link-mini" onClick={() => onNavigate("pinboard")}>Open →</button></div>
          {recent.length === 0
            ? <div className="ov-empty">No mood images yet.</div>
            : <div className="ov-pins">{recent.map((p) => <div key={p.id} className="ov-pin">{p.src ? <img src={p.src} alt="" /> : <div className="thumb-ph" />}</div>)}</div>}
        </div>
        <div className="ov-card">
          <div className="ov-card-head"><span className="font-display">Project</span></div>
          <div className="ov-detail">{client?.address && <div><MapPin size={13} /> {client.address}</div>}<div><Folder size={13} /> {files.length} files · {sharedCount} shared with client</div><div><LayoutGrid size={13} /> {pinboard.length} pinboard images</div><div><Presentation size={13} /> {presentations.length} presentation{presentations.length === 1 ? "" : "s"}</div></div>
          <div className="ov-links">
            <button className="btn btn-ghost" onClick={() => onNavigate("presentations")}><Presentation size={14} /> Presentations</button>
            <button className="btn btn-ghost" onClick={() => onNavigate("files")}><Folder size={14} /> Files</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------- Pinboard / mood board ----------------------------- */
function Pinboard({ clientView, pinboard, onAddImages, onAddUrl, onUpdate, onRemove, onPreview }) {
  const [url, setUrl] = useState("");
  return (
    <div className="pin-wrap">
      {!clientView && (
        <div className="toolbar"><div className="url-add">
          <Pin size={16} className="url-ico" />
          <input className="input url-input" placeholder="Paste an image URL to pin (e.g. a product or inspiration photo)…" value={url}
                 onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { onAddUrl(url); setUrl(""); } }} />
          <button className="btn btn-ghost" onClick={() => { onAddUrl(url); setUrl(""); }} disabled={!url.trim()}><Plus size={15} /> Pin URL</button>
          <label className="btn btn-primary" style={{ cursor: "pointer" }}><Upload size={15} /> Upload<input type="file" multiple accept="image/*" style={{ display: "none" }} onChange={(e) => { onAddImages(e.target.files); e.target.value = ""; }} /></label>
        </div></div>
      )}
      {pinboard.length === 0
        ? <div className="empty" style={{ padding: 70 }}><LayoutGrid size={28} /><p>{clientView ? "Your designer hasn't added mood images yet." : "Build the mood board — upload images or pin from a URL."}</p></div>
        : <div className="masonry">
            {pinboard.map((p) => (
              <figure key={p.id} className="pin">
                {p.src ? <img src={p.src} alt={p.name || ""} loading="lazy" className="zoomable" onClick={() => onPreview && onPreview(p.src)} title="Click to enlarge" /> : <div className="pin-ph"><ImageIcon size={20} /></div>}
                {!clientView && <button className="pin-x" onClick={() => onRemove(p.id)} title="Remove"><X size={13} /></button>}
                {!clientView
                  ? <input className="pin-note" placeholder="Add a note…" value={p.note || ""} onChange={(e) => onUpdate(p.id, { note: e.target.value })} />
                  : (p.note && <figcaption className="pin-cap">{p.note}</figcaption>)}
              </figure>
            ))}
          </div>}
    </div>
  );
}

/* ----------------------------- Files & sharing ----------------------------- */
function FilesShared({ clientView, readOnly = false, files, onAdd, onToggleShared, onRemove }) {
  const visible = clientView ? files.filter((f) => f.sharedWithClient) : files;
  const FileRow = ({ f }) => (
    <div className="file-row">
      {f.type && f.type.startsWith("image/") && f.dataUrl ? <img className="file-thumb" src={f.dataUrl} alt="" /> : <div className="file-thumb ph"><FileText size={16} /></div>}
      <div className="file-info">
        {f.dataUrl ? <a href={f.dataUrl} target="_blank" rel="noreferrer" download={f.name}>{f.name}</a> : <span>{f.name}</span>}
        <span className="file-meta">{f.tooBig ? "too large to store (>4MB)" : `${Math.round((f.size || 0) / 1024)} KB`} · {fmtDate(f.addedAt)}{f.uploadedBy === "client" ? " · from client" : ""}</span>
      </div>
      {!clientView && (
        <>
          <button className={"chip" + (f.sharedWithClient ? " on" : "")} onClick={() => onToggleShared(f.id)} title="Toggle visibility to client">
            {f.sharedWithClient ? <Eye size={12} /> : <EyeOff size={12} />} {f.sharedWithClient ? "Shared" : "Internal"}
          </button>
          <button className="icon-btn" onClick={() => onRemove(f.id)} title="Remove file"><Trash2 size={13} /></button>
        </>
      )}
      {clientView && f.uploadedBy === "client" && <span className="chip on"><Upload size={12} /> You uploaded</span>}
    </div>
  );

  return (
    <div className="files-wrap">
      {!clientView && (
        <div className="toolbar"><div className="url-add" style={{ justifyContent: "flex-start" }}>
          <Folder size={16} className="url-ico" />
          <span style={{ flex: 1, fontSize: 15.5, color: "var(--ink-soft)" }}>Floor plans, contracts, spec sheets, drawings. Mark a file “Shared” to show it on the client dashboard.</span>
          <label className="btn btn-primary" style={{ cursor: "pointer" }}><Upload size={15} /> Upload files<input type="file" multiple style={{ display: "none" }} onChange={(e) => { onAdd(e.target.files, "designer"); e.target.value = ""; }} /></label>
        </div></div>
      )}

      <div className="files-list">
        {clientView && !readOnly && (
          <div className="client-upload">
            <Upload size={20} />
            <div><strong>Share files with your designer</strong><div className="muted" style={{ fontSize: 14.5 }}>Inspiration, documents, anything useful for the project.</div></div>
            <label className="btn btn-primary" style={{ cursor: "pointer" }}>Upload<input type="file" multiple style={{ display: "none" }} onChange={(e) => { onAdd(e.target.files, "client"); e.target.value = ""; }} /></label>
          </div>
        )}
        {visible.length === 0 && !clientView && <div className="empty" style={{ padding: 50 }}><Folder size={26} /><p>No files yet. Upload floor plans, contracts, or specs.</p></div>}
        {visible.length === 0 && clientView && <div className="empty" style={{ padding: 30 }}><p className="muted">No shared files yet.</p></div>}
        {visible.map((f) => <FileRow key={f.id} f={f} />)}
      </div>
    </div>
  );
}

/* ----------------------------- Contractors & trades ----------------------------- */
/* ----------------------------- Finance ----------------------------- */
function Finance({ payments, onPayments, items, discounts, contractors }) {
  const blank = () => mkPayment({ amount: "" });
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState(blank());

  const num = (v) => Number(v) || 0;
  const collected = payments.filter((p) => ["Cleared", "Deposited"].includes(p.status)).reduce((s, p) => s + num(p.amount), 0);
  const collectedCount = payments.filter((p) => ["Cleared", "Deposited"].includes(p.status)).length;
  const pending = payments.filter((p) => p.status === "Pending").reduce((s, p) => s + num(p.amount), 0);
  const estimate = items.reduce((s, i) => s + lineTotal(i, discounts, true), 0) * (1 + TAX_RATE);
  const spentItems = items.reduce((s, i) => s + num(i.actualCost), 0);
  const spentTrades = (contractors || []).reduce((s, c) => s + paidOf(c), 0);
  const spent = spentItems + spentTrades;
  const balanceDue = estimate - collected;

  const save = () => {
    if (form.amount === "" || isNaN(Number(form.amount))) return;
    const rec = { ...form, amount: Number(form.amount) };
    onPayments((prev) => (prev.some((p) => p.id === rec.id) ? prev.map((p) => (p.id === rec.id ? rec : p)) : [rec, ...prev]));
    setForm(blank()); setAdding(false);
  };
  const edit = (p) => { setForm({ ...p }); setAdding(true); };
  const remove = (id) => { if (confirm("Delete this payment?")) onPayments((prev) => prev.filter((p) => p.id !== id)); };

  const Tile = ({ label, value, sub, tone }) => (
    <div className={"fin-tile" + (tone ? " " + tone : "")}>
      <div className="fin-tile-label">{label}</div>
      <div className="fin-tile-val">{money(value)}</div>
      {sub && <div className="fin-tile-sub">{sub}</div>}
    </div>
  );

  return (
    <div className="fin-wrap">
      <div className="fin-tiles">
        <Tile label="Collected from client" value={collected} sub={`${collectedCount} payment${collectedCount === 1 ? "" : "s"} in`} tone="good" />
        <Tile label="Pending / uncleared" value={pending} sub="awaiting deposit or clearing" />
        <Tile label="Project estimate" value={estimate} sub="furniture incl. ~9% tax" />
        <Tile label="Spent to date" value={spent} sub="orders + trades paid" />
        <Tile label={balanceDue >= 0 ? "Balance due from client" : "Collected over estimate"} value={Math.abs(balanceDue)} sub="estimate − collected" tone={balanceDue < 0 ? "good" : ""} />
      </div>

      <div className="toolbar"><div className="url-add" style={{ justifyContent: "space-between" }}>
        <span style={{ flex: 1, fontSize: 14, color: "var(--ink-soft)" }}>Log the client’s checks &amp; payments — track method, reference, what it covers, and whether it has cleared.</span>
        <button className="btn btn-primary" onClick={() => { setForm(blank()); setAdding((a) => !a); }}><Plus size={15} /> Record payment</button>
      </div></div>

      {adding && (
        <div className="fin-form">
          <label className="field"><span>Date</span><input className="input" type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></label>
          <label className="field"><span>Amount</span><input className="input" type="number" placeholder="0.00" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></label>
          <label className="field"><span>Method</span><select className="select" value={form.method} onChange={(e) => setForm({ ...form, method: e.target.value })}>{PAYMENT_METHODS.map((m) => <option key={m}>{m}</option>)}</select></label>
          <label className="field"><span>Reference (check #, conf. #)</span><input className="input" placeholder="e.g. Check #1042" value={form.reference} onChange={(e) => setForm({ ...form, reference: e.target.value })} /></label>
          <label className="field"><span>For</span><select className="select" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>{PAYMENT_CATEGORIES.map((m) => <option key={m}>{m}</option>)}</select></label>
          <label className="field"><span>Status</span><select className="select" value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>{PAYMENT_STATUS.map((m) => <option key={m}>{m}</option>)}</select></label>
          <label className="field field-wide"><span>Notes</span><input className="input" value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></label>
          <div className="fin-form-foot"><button className="btn btn-ghost" onClick={() => { setAdding(false); setForm(blank()); }}>Cancel</button><button className="btn btn-primary" onClick={save}><Check size={15} /> Save payment</button></div>
        </div>
      )}

      <div className="table-wrap">
        <table className="sheet fin-table">
          <thead><tr><th>Date</th><th>Method</th><th>Reference</th><th>For</th><th className="r">Amount</th><th>Status</th><th>Notes</th><th></th></tr></thead>
          <tbody>
            {payments.length === 0 && <tr><td colSpan={8} className="muted" style={{ padding: "18px 12px" }}>No payments logged yet — record the client’s first check above.</td></tr>}
            {payments.map((p) => { const [bg, fg] = PAYMENT_STATUS_COLOR[p.status] || ["#eee", "#444"]; return (
              <tr key={p.id} className="fin-row" onClick={() => edit(p)}>
                <td>{fmtDate(p.date)}</td>
                <td>{p.method}</td>
                <td>{p.reference || "—"}</td>
                <td>{p.category}</td>
                <td className="r num-ro">{money(Number(p.amount) || 0)}</td>
                <td><span className="pay-status" style={{ background: bg, color: fg }}>{p.status}</span></td>
                <td className="muted fin-notes">{p.notes || ""}</td>
                <td className="r"><button className="icon-btn" onClick={(e) => { e.stopPropagation(); remove(p.id); }} title="Delete payment"><Trash2 size={13} /></button></td>
              </tr>
            ); })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ----------------------------- Contractors ----------------------------- */
function Contractors({ clientView, contractors, onContractors }) {
  const [open, setOpen] = useState({});
  const update = (id, patch) => onContractors((prev) => prev.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  const remove = (id) => onContractors((prev) => prev.filter((c) => c.id !== id));
  const add = () => onContractors((prev) => [...prev, mkContractor({ status: "Not started" })]);
  const addPayment = (id) => onContractors((prev) => prev.map((c) => c.id === id ? { ...c, payments: [...(c.payments || []), { id: uid(), amount: "", date: today(), note: "" }] } : c));
  const updatePayment = (id, pid, patch) => onContractors((prev) => prev.map((c) => c.id === id ? { ...c, payments: (c.payments || []).map((p) => p.id === pid ? { ...p, ...patch } : p) } : c));
  const removePayment = (id, pid) => onContractors((prev) => prev.map((c) => c.id === id ? { ...c, payments: (c.payments || []).filter((p) => p.id !== pid) } : c));
  const addInvoices = async (id, fileList) => {
    const files = Array.from(fileList || []);
    const reads = await Promise.all(files.map((f) => new Promise((res) => {
      if (f.size > 4 * 1024 * 1024) { res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: "", tooBig: true, addedAt: today() }); return; }
      const r = new FileReader();
      r.onload = () => res({ id: uid(), name: f.name, type: f.type, size: f.size, dataUrl: r.result, addedAt: today() });
      r.onerror = () => res(null);
      r.readAsDataURL(f);
    })));
    onContractors((prev) => prev.map((c) => c.id === id ? { ...c, invoices: [...(c.invoices || []), ...reads.filter(Boolean)] } : c));
  };
  const removeInvoice = (id, iid) => onContractors((prev) => prev.map((c) => c.id === id ? { ...c, invoices: (c.invoices || []).filter((i) => i.id !== iid) } : c));

  const tContract = contractors.reduce((s, c) => s + (Number(c.contractAmount) || 0), 0);
  const tPaid = contractors.reduce((s, c) => s + paidOf(c), 0);
  const tBal = tContract - tPaid;

  const cols = clientView
    ? ["Company", "Trade", "Scope", "Area", "Status"]
    : ["Company", "Trade", "Scope", "Area", "Contract", "Paid", "Balance", "Payment", "Status", ""];

  return (
    <div>
      {!clientView && (
        <div className="ctr-totals">
          <div className="ctr-tot"><span>Contracted</span><strong className="font-display">{money(tContract)}</strong></div>
          <div className="ctr-tot"><span>Paid to date</span><strong className="font-display" style={{ color: "#356037" }}>{money(tPaid)}</strong></div>
          <div className="ctr-tot"><span>Outstanding balance</span><strong className="font-display" style={{ color: tBal > 0 ? "#A33A2C" : "var(--ink)" }}>{money(tBal)}</strong></div>
        </div>
      )}

      {!clientView && (
        <div className="toolbar"><div className="url-add" style={{ justifyContent: "flex-start" }}>
          <Users size={16} className="url-ico" />
          <span style={{ flex: 1, fontSize: 15.5, color: "var(--ink-soft)" }}>Cabinets, paint, tile, countertops, electrical — track each contractor, what was invoiced, and what's still owed.</span>
          <button className="btn btn-primary" onClick={add}><Plus size={15} /> Add contractor</button>
        </div></div>
      )}

      <div className="table-wrap">
        <table className="sheet">
          <thead><tr>{cols.map((c, i) => <th key={i} className={["Contract", "Paid", "Balance"].includes(c) ? "r" : ""}>{c}</th>)}</tr></thead>
          <tbody>
            {contractors.length === 0 && <tr><td colSpan={cols.length}><div className="empty" style={{ padding: 36 }}><Users size={24} /><p>No contractors yet.{!clientView ? " Add your first trade above." : ""}</p></div></td></tr>}
            {contractors.map((c) => {
              const bal = balanceOf(c), paid = paidOf(c);
              const [pc1, pc2] = PAYSTATUS_COLOR[payStatus(c)];
              const [sc1, sc2] = CSTATUS_COLOR[c.status] || ["#eee", "#444"];
              const isOpen = open[c.id];
              return (
                <React.Fragment key={c.id}>
                  <tr>
                    <td className="name-col"><Cell value={c.company} onChange={(v) => update(c.id, { company: v })} placeholder="Company / person" />
                      {!clientView && <button className="link-mini" onClick={() => setOpen((o) => ({ ...o, [c.id]: !o[c.id] }))}>{isOpen ? "Hide" : "Payments & invoices"} {isOpen ? <ChevronDown size={11} /> : <ChevronRight size={11} />}</button>}
                    </td>
                    <td><select className="cell" value={c.trade} onChange={(e) => update(c.id, { trade: e.target.value })}><option value="">—</option>{TRADES.map((t) => <option key={t}>{t}</option>)}</select></td>
                    <td className="notes-col"><Cell value={c.scope} onChange={(v) => update(c.id, { scope: v })} placeholder="Scope of work" /></td>
                    <td><Cell value={c.room} onChange={(v) => update(c.id, { room: v })} placeholder="—" /></td>
                    {!clientView && <td className="r"><Cell type="number" cls="num" value={c.contractAmount} onChange={(v) => update(c.id, { contractAmount: v })} placeholder="0" /></td>}
                    {!clientView && <td className="r num-ro" style={{ color: "#356037" }}>{money(paid)}</td>}
                    {!clientView && <td className="r num-ro" style={{ color: bal > 0 ? "#A33A2C" : "var(--ink-soft)", fontWeight: 600 }}>{money(bal)}</td>}
                    {!clientView && <td><span className="appr-pill sm" style={{ background: pc1, color: pc2 }}>{payStatus(c)}</span></td>}
                    <td><select className="cell status" style={{ background: sc1, color: sc2 }} value={c.status} onChange={(e) => update(c.id, { status: e.target.value })}>{CONTRACTOR_STATUS.map((s) => <option key={s}>{s}</option>)}</select></td>
                    {!clientView && <td className="act-col"><button className="icon-btn" onClick={() => remove(c.id)} title="Delete"><Trash2 size={13} /></button></td>}
                  </tr>
                  {isOpen && !clientView && (
                    <tr className="ctr-detail"><td colSpan={cols.length}>
                      <div className="ctr-detail-grid">
                        <div className="ctr-pay">
                          <div className="ctr-sub-head">Payments <button className="mini" onClick={() => addPayment(c.id)}><Plus size={12} /> Log payment</button></div>
                          {(c.payments || []).length === 0 && <div className="muted" style={{ fontSize: 14.5 }}>No payments logged.</div>}
                          {(c.payments || []).map((p) => (
                            <div key={p.id} className="pay-row">
                              <input className="input" type="number" placeholder="Amount" value={p.amount} onChange={(e) => updatePayment(c.id, p.id, { amount: e.target.value === "" ? "" : Number(e.target.value) })} style={{ width: 110 }} />
                              <input className="input" type="date" value={p.date} onChange={(e) => updatePayment(c.id, p.id, { date: e.target.value })} style={{ width: 150 }} />
                              <input className="input" placeholder="Note (deposit, progress…)" value={p.note} onChange={(e) => updatePayment(c.id, p.id, { note: e.target.value })} style={{ flex: 1 }} />
                              <button className="icon-btn" onClick={() => removePayment(c.id, p.id)} title="Remove payment"><X size={13} /></button>
                            </div>
                          ))}
                          <div className="pay-sum">Paid {money(paid)} of {money(Number(c.contractAmount) || 0)} · <strong style={{ color: bal > 0 ? "#A33A2C" : "#356037" }}>{bal > 0 ? `${money(bal)} remaining` : "paid in full"}</strong></div>
                          <div className="ctr-contact">
                            <input className="input" placeholder="Phone" value={c.phone} onChange={(e) => update(c.id, { phone: e.target.value })} />
                            <input className="input" placeholder="Email" value={c.email} onChange={(e) => update(c.id, { email: e.target.value })} />
                          </div>
                          <input className="input" placeholder="Notes" value={c.notes} onChange={(e) => update(c.id, { notes: e.target.value })} style={{ marginTop: 6 }} />
                        </div>
                        <div className="ctr-inv">
                          <div className="ctr-sub-head">Invoices <label className="mini upload"><Paperclip size={12} /> Upload<input type="file" multiple accept="image/*,application/pdf" style={{ display: "none" }} onChange={(e) => { addInvoices(c.id, e.target.files); e.target.value = ""; }} /></label></div>
                          {(c.invoices || []).length === 0 && <div className="muted" style={{ fontSize: 14.5 }}>No invoices attached.</div>}
                          <div className="receipt-list">
                            {(c.invoices || []).map((r) => (
                              <div key={r.id} className="receipt">
                                {r.type && r.type.startsWith("image/") && r.dataUrl ? <img src={r.dataUrl} alt="" /> : <div className="receipt-ph"><FileText size={16} /></div>}
                                <div className="receipt-info">
                                  {r.dataUrl ? <a href={r.dataUrl} target="_blank" rel="noreferrer" download={r.name}>{r.name}</a> : <span>{r.name}</span>}
                                  <span className="receipt-meta">{r.tooBig ? "too large to store (>4MB)" : `${Math.round((r.size || 0) / 1024)} KB`}</span>
                                </div>
                                <button className="icon-btn" onClick={() => removeInvoice(c.id, r.id)} title="Remove invoice"><X size={13} /></button>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </td></tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ----------------------------- Client access ----------------------------- */
function ClientAccessPanel({ client, onUpdateClient, onPreview, onClose }) {
  const [copied, setCopied] = useState(false);
  const code = client?.accessCode || "";
  const copy = () => { try { navigator.clipboard && navigator.clipboard.writeText(code); setCopied(true); setTimeout(() => setCopied(false), 1500); } catch (e) {} };
  return (
    <div className="export-modal" onClick={onClose}>
      <div className="access-card" onClick={(e) => e.stopPropagation()}>
        <div className="access-hd">
          <div className="font-display access-title">Client access</div>
          <button className="icon-btn" onClick={onClose} title="Close"><X size={18} /></button>
        </div>
        <p className="access-copy">Give {client?.clientName || "your client"} a read-only login to view their Overview, shared Presentations, shared Files (and upload their own), and the client view of their selections. Internal costs, contractors, and notes stay hidden.</p>
        <label className="access-toggle">
          <input type="checkbox" checked={!!client?.clientAccess} onChange={(e) => onUpdateClient({ clientAccess: e.target.checked })} />
          <span>Allow this client to sign in</span>
        </label>
        {client?.clientAccess && (
          <>
            <div className="access-code-row">
              <div><div className="access-label">Access code</div><div className="access-code font-display">{code}</div></div>
              <div className="access-code-btns">
                <button className="btn btn-ghost" onClick={copy}><Copy size={14} /> {copied ? "Copied" : "Copy"}</button>
                <button className="btn btn-ghost" onClick={() => { if (confirm("Generate a new code? The old one will stop working.")) onUpdateClient({ accessCode: accessCode() }); }}><RefreshCw size={14} /> New code</button>
              </div>
            </div>
            <div className="access-instructions">Tell your client: open the workspace, choose <strong>Client</strong>, and enter code <strong>{code}</strong>.</div>
            <button className="btn btn-primary access-preview" onClick={onPreview}><Eye size={15} /> Open the client’s view</button>
          </>
        )}
        <div className="access-note"><Lock size={12} /> Prototype: codes are stored in this browser and aren’t secure logins. A hosted version adds real per-client accounts and email invites.</div>
      </div>
    </div>
  );
}

/* ----------------------------- Area manager ----------------------------- */
function AreaPanel({ rooms, items, onAdd, onRename, onDelete, onClose }) {
  const [newName, setNewName] = useState("");
  const countOf = (r) => items.filter((i) => i.room === r).length;
  const submit = () => { onAdd(newName); setNewName(""); };
  return (
    <div className="disc-panel">
      <div className="disc-head">
        <div>
          <strong className="font-display">Areas / parts of the house</strong>
          <div className="disc-hint">Add the rooms and zones for this project; they show as categories and in the room dropdowns. Deleting an area moves its items to “Unassigned.” Editing a name updates every item in that area.</div>
        </div>
        <button className="icon-btn" onClick={onClose} title="Close"><X size={16} /></button>
      </div>
      <div className="area-add">
        <input className="input" placeholder="New area — e.g. Wine Cellar, Mudroom, Guest Bath, Pool House" value={newName}
               onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
        <button className="btn btn-primary" onClick={submit} disabled={!newName.trim()}><Plus size={15} /> Add area</button>
      </div>
      <div className="area-list">
        {rooms.map((r) => (
          <div key={r} className="area-row">
            <input className="input area-name" defaultValue={r}
                   onBlur={(e) => onRename(r, e.target.value)}
                   onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }} />
            <span className="area-count">{countOf(r)} {countOf(r) === 1 ? "item" : "items"}</span>
            <button className="icon-btn" onClick={() => onDelete(r)} title="Delete area"><Trash2 size={13} /></button>
          </div>
        ))}
        {rooms.length === 0 && <div className="muted" style={{ fontSize: 14.5 }}>No areas yet — add one above.</div>}
      </div>
    </div>
  );
}

/* ----------------------------- Discount manager ----------------------------- */
const COMMON_RETAILERS = ["Wayfair", "West Elm", "Pottery Barn", "Living Spaces", "Article", "Room & Board", "Amazon",
  "Build.com", "Perigold", "Target", "Birch Lane", "Ruggable", "Rugs USA", "Crate & Barrel", "CB2", "Arhaus",
  "RH", "Anthropologie", "Lulu and Georgia", "McGee & Co", "Four Hands", "Lamps Plus", "Overstock", "World Market"];

function DiscountPanel({ items, discounts, onChange, onRemove, onClose }) {
  const [newName, setNewName] = useState("");
  const [newPct, setNewPct] = useState("");
  const retailers = useMemo(() => {
    const fromItems = items.map((i) => i.retailer).filter(Boolean);
    const fromDisc = Object.keys(discounts || {});
    return Array.from(new Set([...fromItems, ...fromDisc])).sort((a, b) => a.localeCompare(b));
  }, [items, discounts]);
  const itemRetailers = useMemo(() => new Set(items.map((i) => i.retailer).filter(Boolean)), [items]);

  const addRetailer = () => {
    const name = newName.trim();
    if (!name) return;
    onChange(name, newPct === "" ? 0 : Number(newPct));
    setNewName(""); setNewPct("");
  };

  return (
    <div className="disc-panel">
      <div className="disc-head">
        <div>
          <strong className="font-display">Trade discounts by retailer</strong>
          <div className="disc-hint">Set the % off she gets at each store. Every item from that retailer uses it automatically — override a single item in its row or card. Hidden from clients.</div>
        </div>
        <button className="icon-btn" onClick={onClose} title="Close"><X size={16} /></button>
      </div>

      <div className="disc-grid">
        {retailers.map((r) => (
          <div key={r} className="disc-row">
            <span className="disc-name">{r}</span>
            <div className="disc-input-wrap">
              <input className="input disc-pct" type="number" min="0" max="100" placeholder="0"
                     value={discounts[r] ?? ""}
                     onChange={(e) => onChange(r, e.target.value === "" ? "" : Number(e.target.value))} />
              <span className="disc-sign">%</span>
            </div>
            {!itemRetailers.has(r) && <button className="icon-btn disc-del" onClick={() => onRemove(r)} title="Remove retailer"><X size={14} /></button>}
          </div>
        ))}
        {retailers.length === 0 && <div className="muted" style={{ gridColumn: "1/-1" }}>No retailers yet — add one below or paste a product link.</div>}
      </div>

      <div className="disc-add">
        <div className="disc-add-label">Add a retailer</div>
        <div className="disc-add-row">
          <input className="input" list="dl-retailers" placeholder="Retailer name (e.g. Crate & Barrel)" value={newName}
                 onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addRetailer()} />
          <div className="disc-input-wrap">
            <input className="input disc-pct" type="number" min="0" max="100" placeholder="0" value={newPct}
                   onChange={(e) => setNewPct(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addRetailer()} />
            <span className="disc-sign">%</span>
          </div>
          <button className="btn btn-primary" onClick={addRetailer} disabled={!newName.trim()}><Plus size={15} /> Add</button>
        </div>
        <datalist id="dl-retailers">{COMMON_RETAILERS.map((r) => <option key={r} value={r} />)}</datalist>
      </div>
    </div>
  );
}

/* ----------------------------- Spreadsheet list view ----------------------------- */
function SheetThumb({ it, onPreview }) {
  const [err, setErr] = useState(false);
  return (
    <div className="thumb-sm">
      {it.imageUrl && !err
        ? <img src={it.imageUrl} alt="" className="zoomable" onError={() => setErr(true)} onClick={() => onPreview && onPreview(it.imageUrl)} title="Click to enlarge" />
        : <div className="thumb-sm-ph"><ImageIcon size={13} /></div>}
    </div>
  );
}

const Cell = ({ value, onChange, placeholder, type = "text", cls = "" }) => (
  <input className={"cell " + cls} type={type} value={value ?? ""} placeholder={placeholder}
    onChange={(e) => onChange(type === "number" ? (e.target.value === "" ? "" : Number(e.target.value)) : e.target.value)} />
);

function ScheduleTable({ grouped, clientView, refreshing, rd, scanning, tracking, selectable, selected, onToggleSel, onScan, onTrack, onPreview, onAddReceipts, onRemoveReceipt, onUpdate, onRemove, onRefresh }) {
  const cols = clientView
    ? ["", "Room", "Item", "Source", "Item Name / SKU", "Finish / Selection", "Qty", "Price Each", "Line Total", "Status", "Approval", "Stock", "Lead / ETA", ""]
    : ["", "Room", "Item", "Source", "Item Name / SKU", "Finish / Selection", "Qty", "Price Each", "Disc %", "Line Total", "Est. Tax", "Status", "Approval", "Stock", "Lead / ETA", "Tracking", "Actual", "Last", "Notes", ""];
  const ltIdx = cols.indexOf("Line Total");

  return (
    <div className="table-wrap">
      <table className="sheet">
        <thead><tr>{cols.map((c, i) => <th key={i} className={["Qty", "Price Each", "Disc %", "Line Total", "Est. Tax", "Actual"].includes(c) ? "r" : ""}>{c}</th>)}</tr></thead>
        <tbody>
          {Object.entries(grouped).map(([room, list]) => {
            const sub = list.reduce((s, i) => s + lineTotal(i, rd, !clientView), 0);
            return (
              <React.Fragment key={room}>
                <tr className="grp"><td colSpan={cols.length}><span className="font-display">{room}</span><em>{list.length} {list.length === 1 ? "item" : "items"}</em></td></tr>
                {list.length === 0 && <tr className="area-empty-row"><td colSpan={cols.length}>No items in this area yet — paste a product link or add one manually.</td></tr>}
                {list.map((it) => {
                  const price = effPrice(it);
                  const overridden = it.priceOverride !== "" && it.priceOverride !== null;
                  const oos = (it.availability || "").toLowerCase().includes("out");
                  const [sc1, sc2] = STATUS_COLOR[it.status] || ["#eee", "#444"];
                  const [ac1, ac2] = APPROVAL_COLOR[it.approval.state] || APPROVAL_COLOR.Pending;
                  return (
                    <tr key={it.id} className={selectable && selected?.has(it.id) ? "row-sel" : ""}>
                      <td className="thumb-col">
                        {selectable && <button className={"sel-box tbl" + (selected?.has(it.id) ? " on" : "")} onClick={() => onToggleSel(it.id)} title="Select for export">{selected?.has(it.id) ? <CheckSquare size={14} /> : <Square size={14} />}</button>}
                        <SheetThumb it={it} onPreview={onPreview} />
                      </td>
                      <td className="room-col"><Cell value={it.room} onChange={(v) => onUpdate(it.id, { room: v })} /></td>
                      <td className="type-col"><Cell value={it.itemType} onChange={(v) => onUpdate(it.id, { itemType: v })} /></td>
                      <td className="src-col"><Cell value={it.retailer} onChange={(v) => onUpdate(it.id, { retailer: v })} /></td>
                      <td className="name-col">
                        <Cell value={it.name} onChange={(v) => onUpdate(it.id, { name: v })} placeholder="Product name" />
                        {it.sku && <div className="sku-sub">#{it.sku}</div>}
                      </td>
                      <td className="finish-col"><Cell value={it.finish} onChange={(v) => onUpdate(it.id, { finish: v })} />{it.dimensions && <div className="dim-sub">{it.dimensions}</div>}</td>
                      <td className="r"><Cell type="number" cls="num qty" value={it.qty} onChange={(v) => onUpdate(it.id, { qty: v })} /></td>
                      <td className="r">
                        <div className="price-cell">
                          <Cell type="number" cls="num" value={overridden ? it.priceOverride : it.priceAuto}
                                onChange={(v) => onUpdate(it.id, { priceOverride: v })} />
                          {overridden && <button className="revert" title={`Auto: ${money(it.priceAuto)} — click to revert`} onClick={() => onUpdate(it.id, { priceOverride: "" })}><RotateCcw size={11} /></button>}
                        </div>
                        {it.lastPriceChange && (
                          <div className={"chg " + (it.lastPriceChange.new > it.lastPriceChange.old ? "up" : "down")} title={`Was ${money(it.lastPriceChange.old)}`}>
                            {it.lastPriceChange.new > it.lastPriceChange.old ? <ArrowUp size={9} /> : <ArrowDown size={9} />}{money(Math.abs(it.lastPriceChange.new - it.lastPriceChange.old))}
                          </div>
                        )}
                      </td>
                      {!clientView && (
                        <td className="r disc-col">
                          <Cell type="number" cls="num disc" value={it.discountPct}
                                placeholder={String(Number(rd[it.retailer]) || 0)}
                                onChange={(v) => onUpdate(it.id, { discountPct: v })} />
                          {discPct(it, rd) > 0 && <div className="net-sub">{money(netEach(it, rd))} ea</div>}
                        </td>
                      )}
                      <td className="r num-ro">{money(lineTotal(it, rd, !clientView))}</td>
                      {!clientView && <td className="r num-ro tax">{money(lineTotal(it, rd, true) * TAX_RATE)}</td>}
                      <td>
                        <select className="cell status" style={{ background: sc1, color: sc2 }} value={it.status}
                                onChange={(e) => onUpdate(it.id, { status: e.target.value })}>
                          {STATUSES.map((s) => <option key={s}>{s}</option>)}
                        </select>
                      </td>
                      <td><span className="appr-pill sm" style={{ background: ac1, color: ac2 }}>{it.approval.state}</span></td>
                      <td>{oos ? <span className="avail oos sm"><AlertTriangle size={10} /> Out</span> : <span className="avail sm">{(it.availability || "—").replace("In stock", "In stock")}</span>}</td>
                      <td className="lead-col">
                        <Cell value={it.leadTime} onChange={(v) => onUpdate(it.id, { leadTime: v })} placeholder="—" />
                        {it.eta && <div className="track-sub">ETA {fmtDate(it.eta)}</div>}
                      </td>
                      {!clientView && (
                        <td className="track-col">
                          <div className="track-cell">
                            <Cell value={it.trackingNumber} onChange={(v) => onUpdate(it.id, { trackingNumber: v })} placeholder="tracking #" />
                            <button className="icon-btn" onClick={() => onTrack(it)} disabled={tracking === it.id} title="Fetch shipment status">{tracking === it.id ? <Loader2 size={12} className="spin" /> : <Truck size={12} />}</button>
                            {it.trackingUrl && <a className="icon-btn" href={it.trackingUrl} target="_blank" rel="noreferrer" title="Open tracking page"><ExternalLink size={12} /></a>}
                          </div>
                          {it.shipStatus && <div className="track-sub ship">{it.shipStatus}</div>}
                        </td>
                      )}
                      {!clientView && <td className="r"><Cell type="number" cls="num" value={it.actualCost} onChange={(v) => onUpdate(it.id, { actualCost: v })} placeholder="—" /></td>}
                      {!clientView && <td className="last-col">{fmtDate(it.lastChecked)}</td>}
                      {!clientView && <td className="notes-col"><Cell value={it.notes} onChange={(v) => onUpdate(it.id, { notes: v })} placeholder="—" /></td>}
                      <td className="act-col">
                        {clientView ? (
                          <div className="row-actions">
                            <button className="icon-btn ok" title="Approve" onClick={() => onUpdate(it.id, { approval: { state: "Approved", date: today() }, status: "Client approved" })}><Check size={13} /></button>
                            <button className="icon-btn bad" title="Reject" onClick={() => onUpdate(it.id, { approval: { state: "Rejected", date: today() } })}><X size={13} /></button>
                          </div>
                        ) : (
                          <div className="row-actions">
                            {it.url && <button className="icon-btn" onClick={() => onRefresh(it)} disabled={refreshing === it.id} title="Re-check price & stock">{refreshing === it.id ? <Loader2 size={13} className="spin" /> : <RefreshCw size={13} />}</button>}
                            <button className="icon-btn" onClick={() => onScan(it)} disabled={scanning === it.id} title="Scan Gmail for order & tracking">{scanning === it.id ? <Loader2 size={13} className="spin" /> : <Inbox size={13} />}</button>
                            <label className="icon-btn" title="Attach receipt / invoice" style={{ position: "relative" }}>
                              <Paperclip size={13} />
                              {(it.receipts || []).length > 0 && <span className="rcpt-badge">{it.receipts.length}</span>}
                              <input type="file" multiple accept="image/*,application/pdf" style={{ display: "none" }} onChange={(e) => { onAddReceipts(it.id, e.target.files); e.target.value = ""; }} />
                            </label>
                            {it.url && <a className="icon-btn" href={it.url} target="_blank" rel="noreferrer" title="Open page"><ExternalLink size={13} /></a>}
                            <button className="icon-btn" onClick={() => onRemove(it.id)} title="Delete"><Trash2 size={13} /></button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
                {list.length > 0 && (
                <tr className="subtotal">
                  <td colSpan={ltIdx} className="r">{room} subtotal</td>
                  <td className="r num-ro">{money(sub)}</td>
                  {!clientView && <td className="r num-ro">{money(sub * TAX_RATE)}</td>}
                  <td colSpan={cols.length - ltIdx - (clientView ? 1 : 2)}></td>
                </tr>
                )}
              </React.Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/* ----------------------------- Totals ----------------------------- */
function Totals({ label, sub, tax, total, count, savings, clientView }) {
  return (
    <div className="totals-inner">
      <div className="tot-item"><span>{label}{!clientView ? " (after discount)" : ""}</span><strong className="font-display">{money(total)}</strong></div>
      <div className="tot-break">
        <span>{count} items</span><span>Subtotal {money(sub)}</span><span>Est. tax {money(tax)}</span>
        {!clientView && savings > 0 && <span className="savings">Discount savings {money(savings)}</span>}
      </div>
    </div>
  );
}

/* ----------------------------- Item card ----------------------------- */
function ItemCard({ it, clientView, rd = {}, refreshing, scanning, tracking, selectable, sel, onSel, onScan, onTrack, onAddReceipts, onRemoveReceipt, onComment, onDecide, onPreview, readOnly = false, onUpdate, onRemove, onRefresh }) {
  const [open, setOpen] = useState(false);
  const [thread, setThread] = useState(false);
  const [draft, setDraft] = useState("");
  const [imgErr, setImgErr] = useState(false);
  const comments = it.comments || [];
  const sendComment = () => { if (!draft.trim()) return; onComment(clientView ? "client" : "designer", draft); setDraft(""); };
  const price = effPrice(it);
  const disc = discPct(it, rd);
  const net = netEach(it, rd);
  const overridden = it.priceOverride !== "" && it.priceOverride !== null && Number(it.priceOverride) !== Number(it.priceAuto);
  const oos = (it.availability || "").toLowerCase().includes("out");
  const [sc1, sc2] = STATUS_COLOR[it.status] || ["#eee", "#444"];

  return (
    <div className={"card" + (sel ? " sel" : "")}>
      <div className="card-main">
        {/* photo */}
        <div className="thumb">
          {selectable && <button className={"sel-box" + (sel ? " on" : "")} onClick={onSel} title="Select for export">{sel ? <CheckSquare size={15} /> : <Square size={15} />}</button>}
          {it.imageUrl && !imgErr
            ? <img src={it.imageUrl} alt="" className="zoomable" onError={() => setImgErr(true)} onClick={() => onPreview && onPreview(it.imageUrl)} title="Click to enlarge" />
            : <div className="thumb-ph"><ImageIcon size={18} /><span>{it.itemType || "Item"}</span></div>}
        </div>

        {/* core */}
        <div className="card-body">
          <div className="card-top">
            <div className="retailer">{it.retailer || "—"}</div>
            <div className="status-wrap">
              {!clientView ? (
                <select className="status-pill" style={{ background: sc1, color: sc2 }}
                        value={it.status} onChange={(e) => onUpdate({ status: e.target.value })}>
                  {STATUSES.map((s) => <option key={s}>{s}</option>)}
                </select>
              ) : (
                <span className="status-pill" style={{ background: sc1, color: sc2 }}>{it.status}</span>
              )}
            </div>
          </div>

          <div className="name">{it.name || <em className="muted">Unnamed item</em>}</div>
          <div className="meta">
            {it.finish && <span>{it.finish}</span>}
            {it.dimensions && <span>{it.dimensions}</span>}
            {it.sku && <span className="sku">#{it.sku}</span>}
          </div>

          <div className="card-foot">
            <div className="price-block">
              <span className="qty">{it.qty} ×</span>
              <span className="price" style={!clientView && disc > 0 ? { textDecoration: "line-through", color: "var(--ink-soft)", fontWeight: 500, fontSize: 15.5 } : undefined}>{money(price)}</span>
              {!clientView && disc > 0 && <span className="tag tag-disc">−{disc}%</span>}
              {!clientView && disc > 0 && <span className="price">{money(net)}</span>}
              {overridden && <span className="tag tag-edit">override</span>}
              {it.lastPriceChange && (
                <span className={"tag " + (it.lastPriceChange.new > it.lastPriceChange.old ? "tag-up" : "tag-down")}
                      title={`Was ${money(it.lastPriceChange.old)} on last check`}>
                  {it.lastPriceChange.new > it.lastPriceChange.old ? <ArrowUp size={11} /> : <ArrowDown size={11} />}
                  {money(Math.abs(it.lastPriceChange.new - it.lastPriceChange.old))}
                </span>
              )}
              <span className="line-tot">= {money(lineTotal(it, rd, !clientView))}</span>
            </div>
            <div className="avail-block">
              {oos ? <span className="avail oos"><AlertTriangle size={11} /> Out of stock</span>
                   : <span className="avail">{it.availability}</span>}
              <span className="checked">checked {fmtDate(it.lastChecked)}</span>
            </div>
          </div>

          {!clientView && (it.trackingNumber || it.carrier || it.eta || it.shipStatus || (it.receipts && it.receipts.length > 0)) && (
            <div className="track-strip">
              {(it.carrier || it.trackingNumber) && (
                it.trackingUrl
                  ? <a className="track-pill" href={it.trackingUrl} target="_blank" rel="noreferrer"><Truck size={11} /> {it.carrier || "Tracking"}{it.trackingNumber ? ` · ${it.trackingNumber}` : ""}</a>
                  : <span className="track-pill"><Truck size={11} /> {it.carrier || "Tracking"}{it.trackingNumber ? ` · ${it.trackingNumber}` : ""}</span>
              )}
              {it.shipStatus && <span className="track-eta">{it.shipStatus}</span>}
              {it.eta && <span className="track-eta">ETA {fmtDate(it.eta)}</span>}
              {it.payMethod && <span className="track-pill pay"><CreditCard size={11} /> {it.payMethod}{it.payRef ? ` · ${it.payRef}` : ""}</span>}
              {it.receipts && it.receipts.length > 0 && <span className="track-pill"><Paperclip size={11} /> {it.receipts.length} {it.receipts.length === 1 ? "receipt" : "receipts"}</span>}
            </div>
          )}

          {/* approval row (client can act in client view) */}
          <div className="approval-row">
            <ApprovalPill approval={it.approval} />
            {clientView ? (
              <div className="approve-btns">
                <button className="link-mini cmt-link" onClick={() => setThread((o) => !o)}><MessageSquare size={12} /> {comments.length ? comments.length : "Note"}</button>
                {!readOnly && <button className={"mini ok" + (it.approval.state === "Approved" ? " active" : "")} onClick={() => onDecide("approve")}><Check size={13} /> Approve</button>}
                {!readOnly && <button className={"mini bad" + (it.approval.state === "Rejected" ? " active" : "")} onClick={() => { setThread(true); onDecide("reject"); }}><X size={13} /> Request changes</button>}
              </div>
            ) : (
              <div className="row-actions">
                <button className="link-mini cmt-link" onClick={() => setThread((o) => !o)} title="Comments"><MessageSquare size={14} />{comments.length > 0 && <span className="cmt-badge">{comments.length}</span>}</button>
                {it.url && (
                  <button className="icon-btn" onClick={onRefresh} disabled={refreshing} title="Re-check price & stock">
                    {refreshing ? <Loader2 size={14} className="spin" /> : <RefreshCw size={14} />}
                  </button>
                )}
                <button className="icon-btn" onClick={onScan} disabled={scanning} title="Scan Gmail for order & tracking">
                  {scanning ? <Loader2 size={14} className="spin" /> : <Inbox size={14} />}
                </button>
                {it.url && <a className="icon-btn" href={it.url} target="_blank" rel="noreferrer" title="Open product page"><ExternalLink size={14} /></a>}
                <button className="icon-btn" onClick={() => setOpen((o) => !o)} title="Edit">
                  {open ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                </button>
              </div>
            )}
          </div>

          {/* comment thread (both sides) */}
          {thread && (
            <div className="thread">
              {comments.length === 0 && <div className="thread-empty">{clientView ? "Leave a note or question for your designer." : "No comments yet."}</div>}
              {comments.map((c) => (
                <div key={c.id} className={"cmt " + c.by}>
                  <span className="cmt-by">{c.by === "client" ? "Client" : "Designer"}</span>
                  <span className="cmt-text">{c.text}</span>
                  <span className="cmt-when">{fmtWhen(c.date)}</span>
                </div>
              ))}
              {!readOnly && <div className="thread-input">
                <input className="input" placeholder={clientView ? "Add a note for your designer…" : "Reply to your client…"} value={draft}
                       onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => e.key === "Enter" && sendComment()} />
                <button className="btn btn-primary" onClick={sendComment} disabled={!draft.trim()}>Send</button>
              </div>}
            </div>
          )}
        </div>
      </div>

      {/* expanded editor (internal only) */}
      {open && !clientView && (
        <div className="editor">
          <div className="field-wide"><div className="editor-section" style={{ marginTop: 0 }}>Product photo</div><ImageInput value={it.imageUrl} onChange={(v) => onUpdate({ imageUrl: v })} /></div>
          <Field label="Room"><input className="input" value={it.room} onChange={(e) => onUpdate({ room: e.target.value })} /></Field>
          <Field label="Item type"><input className="input" value={it.itemType} onChange={(e) => onUpdate({ itemType: e.target.value })} /></Field>
          <Field label="Finish / Selection"><input className="input" value={it.finish} onChange={(e) => onUpdate({ finish: e.target.value })} /></Field>
          <Field label="Qty"><input className="input" type="number" min="0" value={it.qty} onChange={(e) => onUpdate({ qty: Number(e.target.value) })} /></Field>
          <Field label="Auto price"><input className="input" value={it.priceAuto} onChange={(e) => onUpdate({ priceAuto: Number(e.target.value) })} /></Field>
          <Field label="Manual override"><input className="input" placeholder="(uses auto)" value={it.priceOverride} onChange={(e) => onUpdate({ priceOverride: e.target.value })} /></Field>
          <Field label={`Discount % (${it.retailer || "retailer"} default ${Number(rd[it.retailer]) || 0}%)`}><input className="input" type="number" placeholder={String(Number(rd[it.retailer]) || 0)} value={it.discountPct} onChange={(e) => onUpdate({ discountPct: e.target.value === "" ? "" : Number(e.target.value) })} /></Field>
          <Field label="Lead time / ETA"><input className="input" value={it.leadTime} onChange={(e) => onUpdate({ leadTime: e.target.value })} /></Field>
          <Field label="Actual cost"><input className="input" placeholder="(when paid)" value={it.actualCost} onChange={(e) => onUpdate({ actualCost: e.target.value })} /></Field>

          <div className="editor-section">Order tracking
            {it.lastScanned && <span className="scan-note">{it.scanResult === "found" ? "✓ matched an inbox email" : "no matching email found"} · scanned {fmtDate(it.lastScanned)}</span>}
            <button className="mini" onClick={onScan} disabled={scanning} style={{ marginLeft: "auto" }}>{scanning ? <Loader2 size={12} className="spin" /> : <Inbox size={12} />} Scan inbox</button>
          </div>
          <Field label="Order #"><input className="input" value={it.orderNumber} onChange={(e) => onUpdate({ orderNumber: e.target.value })} /></Field>
          <Field label="Order date"><input className="input" type="date" value={it.orderDate} onChange={(e) => onUpdate({ orderDate: e.target.value })} /></Field>
          <Field label={"Payment method" + (ORDERED_STATUSES.includes(it.status) ? "" : " (once ordered)")}>
            <select className="select" value={it.payMethod} onChange={(e) => onUpdate({ payMethod: e.target.value })}>
              <option value="">—</option>
              {PAY_METHODS.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
          <Field label="Payment ref (check #, last 4…)"><input className="input" placeholder="e.g. Check #1042" value={it.payRef} onChange={(e) => onUpdate({ payRef: e.target.value })} /></Field>
          <Field label="Carrier"><input className="input" placeholder="UPS / FedEx / freight" value={it.carrier} onChange={(e) => onUpdate({ carrier: e.target.value })} /></Field>
          <Field label="Tracking #"><input className="input" value={it.trackingNumber} onChange={(e) => onUpdate({ trackingNumber: e.target.value })} /></Field>
          <Field label="Tracking URL" wide>
            <div style={{ display: "flex", gap: 6 }}>
              <input className="input" placeholder="https://…  (paste the carrier's tracking link)" value={it.trackingUrl} onChange={(e) => onUpdate({ trackingUrl: e.target.value })} />
              <button className="btn btn-ghost" onClick={onTrack} disabled={tracking} style={{ whiteSpace: "nowrap" }}>{tracking ? <Loader2 size={14} className="spin" /> : <Truck size={14} />} Fetch status</button>
            </div>
          </Field>
          {(it.shipStatus || it.shipChecked) && <div className="field-wide ship-note">{it.shipStatus ? <><Truck size={12} /> {it.shipStatus}</> : <>No live status found</>}{it.shipChecked ? ` · checked ${fmtDate(it.shipChecked)}` : ""}</div>}
          <Field label="Est. delivery"><input className="input" type="date" value={it.eta} onChange={(e) => onUpdate({ eta: e.target.value })} /></Field>

          <div className="editor-section">Receipts &amp; invoices
            <label className="mini upload" style={{ marginLeft: "auto" }}>
              <Paperclip size={12} /> Upload
              <input type="file" multiple accept="image/*,application/pdf" style={{ display: "none" }}
                     onChange={(e) => { onAddReceipts(e.target.files); e.target.value = ""; }} />
            </label>
          </div>
          {(it.receipts || []).length === 0 && <div className="field-wide muted" style={{ fontSize: 14.5 }}>No receipts attached yet.</div>}
          {(it.receipts || []).length > 0 && (
            <div className="field-wide receipt-list">
              {it.receipts.map((r) => (
                <div key={r.id} className="receipt">
                  {r.type && r.type.startsWith("image/") && r.dataUrl
                    ? <img src={r.dataUrl} alt="" />
                    : <div className="receipt-ph"><FileText size={16} /></div>}
                  <div className="receipt-info">
                    {r.dataUrl ? <a href={r.dataUrl} target="_blank" rel="noreferrer" download={r.name}>{r.name}</a> : <span>{r.name}</span>}
                    <span className="receipt-meta">{r.tooBig ? "too large to store (>4MB)" : `${Math.round((r.size || 0) / 1024)} KB · ${fmtDate(r.addedAt)}`}</span>
                  </div>
                  <button className="icon-btn" onClick={() => onRemoveReceipt(r.id)} title="Remove"><X size={13} /></button>
                </div>
              ))}
            </div>
          )}

          <Field label="Product URL" wide><input className="input" value={it.url} onChange={(e) => onUpdate({ url: e.target.value })} /></Field>
          <Field label="Notes" wide><input className="input" value={it.notes} onChange={(e) => onUpdate({ notes: e.target.value })} /></Field>
          <div className="editor-foot">
            <button className="btn btn-danger" onClick={onRemove}><Trash2 size={14} /> Delete item</button>
          </div>
        </div>
      )}
    </div>
  );
}

function ApprovalPill({ approval }) {
  const [bg, fg] = APPROVAL_COLOR[approval.state] || APPROVAL_COLOR.Pending;
  return (
    <span className="appr-pill" style={{ background: bg, color: fg }}>
      {approval.state}{approval.date ? ` · ${fmtDate(approval.date)}` : ""}
    </span>
  );
}

function Field({ label, children, wide }) {
  return <label className={"field" + (wide ? " field-wide" : "")}><span>{label}</span>{children}</label>;
}

/* ----------------------------- Draft panel ----------------------------- */
/* ----------------------------- Image input (upload / drag-drop / paste / URL) ----------------------------- */
function ImageInput({ value, onChange }) {
  const [drag, setDrag] = useState(false);
  const [url, setUrl] = useState("");
  const readFile = (file) => {
    if (!file) return;
    if (file.size > 4 * 1024 * 1024) { alert("That image is larger than 4MB — try a smaller file or paste an image URL instead."); return; }
    const r = new FileReader(); r.onload = () => onChange(r.result); r.readAsDataURL(file);
  };
  return (
    <div className={"imgin" + (drag ? " drag" : "")} tabIndex={0}
         onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
         onDrop={(e) => { e.preventDefault(); setDrag(false); const f = Array.from(e.dataTransfer.files || []).find((x) => x.type.startsWith("image/")); if (f) readFile(f); }}
         onPaste={(e) => { const img = Array.from(e.clipboardData?.items || []).find((it) => it.type.startsWith("image/")); if (img) { readFile(img.getAsFile()); e.preventDefault(); return; } const t = e.clipboardData?.getData("text"); if (t && /^https?:\/\//.test(t.trim())) { onChange(t.trim()); e.preventDefault(); } }}>
      {value
        ? <div className="imgin-has"><img src={value} alt="" /><button className="imgin-clear" onClick={() => onChange("")} title="Remove image"><X size={13} /></button></div>
        : <div className="imgin-empty"><ImageIcon size={22} /><span>Drag &amp; drop or paste an image</span><label className="imgin-browse">browse files<input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { readFile(e.target.files?.[0]); e.target.value = ""; }} /></label></div>}
      <div className="imgin-url"><input className="input" placeholder="or paste an image URL" value={url} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && url.trim()) { onChange(url.trim()); setUrl(""); } }} /></div>
    </div>
  );
}

function DraftPanel({ draft, setDraft, fetching, rooms, types, onCommit, onCancel, onPreview }) {
  const set = (patch) => setDraft((d) => ({ ...d, ...patch }));
  useEffect(() => {
    const onPaste = (e) => {
      const img = Array.from(e.clipboardData?.items || []).find((it) => it.type && it.type.startsWith("image/"));
      if (!img) return;
      const f = img.getAsFile(); if (!f || f.size > 4 * 1024 * 1024) return;
      const r = new FileReader(); r.onload = () => set({ imageUrl: r.result }); r.readAsDataURL(f);
    };
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, []);
  return (
    <div className="draft">
      <div className="draft-head">
        <strong className="font-display">{draft.url ? "Review fetched item" : "Add item manually"}</strong>
        <button className="icon-btn" onClick={onCancel} title="Close"><X size={16} /></button>
      </div>

      {fetching && <div className="draft-status"><Loader2 size={14} className="spin" /> Fetching product details…</div>}
      {draft._fetched && <div className="draft-status ok"><Check size={14} /> Auto-filled. Review and adjust anything below.</div>}
      {draft._fetchFailed && <div className="draft-status warn"><AlertTriangle size={14} /> Couldn’t fetch this site automatically — fill the details in by hand.</div>}

      <div className="draft-grid">
        <div className="draft-photo">
          <div className="draft-photo-label">Product photo {draft.imageUrl ? <span className="link-mini" onClick={() => onPreview && onPreview(draft.imageUrl)} style={{ cursor: "pointer" }}>· preview</span> : <span className="muted">· couldn’t fetch? add your own</span>}</div>
          <ImageInput value={draft.imageUrl} onChange={(v) => set({ imageUrl: v })} />
        </div>
        <div className="draft-fields">
          <Field label="Product name" wide><input className="input" value={draft.name} onChange={(e) => set({ name: e.target.value })} /></Field>
          <Field label="Retailer"><input className="input" value={draft.retailer} onChange={(e) => set({ retailer: e.target.value })} /></Field>
          <Field label="SKU / Model"><input className="input" value={draft.sku} onChange={(e) => set({ sku: e.target.value })} /></Field>
          <Field label="Room"><input className="input" list="dl-rooms" value={draft.room} onChange={(e) => set({ room: e.target.value })} /></Field>
          <Field label="Item type"><input className="input" list="dl-types" value={draft.itemType} onChange={(e) => set({ itemType: e.target.value })} /></Field>
          <Field label="Finish / Selection"><input className="input" value={draft.finish} onChange={(e) => set({ finish: e.target.value })} /></Field>
          <Field label="Dimensions"><input className="input" value={draft.dimensions} onChange={(e) => set({ dimensions: e.target.value })} /></Field>
          <Field label="Price each"><input className="input" type="number" value={draft.priceAuto} onChange={(e) => set({ priceAuto: Number(e.target.value) })} /></Field>
          <Field label="Qty"><input className="input" type="number" min="1" value={draft.qty} onChange={(e) => set({ qty: Number(e.target.value) })} /></Field>
          <Field label="Availability"><input className="input" value={draft.availability} onChange={(e) => set({ availability: e.target.value })} /></Field>
          <Field label="Status">
            <select className="select" value={draft.status} onChange={(e) => set({ status: e.target.value })}>
              {STATUSES.map((s) => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <Field label="Product URL" wide><input className="input" value={draft.url} onChange={(e) => set({ url: e.target.value })} /></Field>
        </div>
      </div>

      <datalist id="dl-rooms">{rooms.map((r) => <option key={r} value={r} />)}</datalist>
      <datalist id="dl-types">{types.map((t) => <option key={t} value={t} />)}</datalist>

      <div className="draft-foot">
        <button className="btn btn-ghost" onClick={onCancel}>Cancel</button>
        <button className="btn btn-primary" onClick={onCommit} disabled={fetching}><Check size={15} /> Add to schedule</button>
      </div>
    </div>
  );
}

/* ----------------------------- styles ----------------------------- */
function Style() {
  return (
    <style>{`
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&family=Jost:wght@300;400;500;600&display=swap');
:root{
  --paper:#FAF8F2; --card:#FFFFFF; --card-2:#FCFAF5; --ink:#2B2A2A; --ink-soft:#615D55; --ink-faint:#938E83;
  --line:#ECE7DD; --line-2:#DED8CC; --clay:#33323A; --clay-soft:#EDEBE6; --sage:#5E6E54; --gold:#8A6B2C;
  --sand:#F4F1E9;
}
*{box-sizing:border-box}
.app{font-family:'Jost',sans-serif;color:var(--ink);font-size:16.5px;line-height:1.5;font-weight:400;-webkit-font-smoothing:antialiased}
.font-display{font-family:'Cormorant Garamond',Georgia,serif;font-weight:600;letter-spacing:.005em}
.spin{animation:sp 1s linear infinite}@keyframes sp{to{transform:rotate(360deg)}}
.muted{color:var(--ink-soft)}
em{font-style:normal}

/* header */
.hdr{padding:24px 30px 0;background:var(--paper)}
.hdr-row{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap}
.eyebrow{font-size:12.5px;letter-spacing:.22em;text-transform:uppercase;color:var(--ink-soft);font-weight:500}
.title{font-size:44.5px;font-weight:600;margin:2px 0 0;letter-spacing:.005em;line-height:1.05}
.hdr-actions{display:flex;gap:8px;align-items:center}
.totals{margin-top:18px}
.totals-inner{display:inline-flex;flex-direction:column;gap:5px;background:var(--card);border:1px solid var(--line);
  border-radius:12px;padding:13px 20px}
.tot-item{display:flex;align-items:baseline;gap:14px}
.tot-item span{font-size:13.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-soft)}
.tot-item strong{font-size:29.7px;font-weight:500}
.tot-break{display:flex;gap:18px;font-size:14.5px;color:var(--ink-soft)}
.banner{margin-top:14px;display:flex;align-items:center;gap:8px;background:var(--clay-soft);color:#7d4a2e;
  padding:9px 14px;border-radius:9px;font-size:15px;font-weight:500}

/* toolbar */
.toolbar{padding:18px 30px 0}
.url-add{display:flex;align-items:center;gap:8px;background:var(--card);border:1px solid var(--line);
  border-radius:10px;padding:7px 8px 7px 14px}
.url-ico{color:var(--ink-soft);flex:none}
.url-input{border:none!important;background:transparent!important;flex:1;padding:6px 0!important}
.url-input:focus{outline:none}

/* filters */
.filters{display:flex;gap:10px;padding:16px 26px;flex-wrap:wrap}
.search-wrap{position:relative;flex:1;min-width:200px}
.search-ico{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--ink-soft)}
.search-wrap .input{padding-left:32px!important}
.input{width:100%;border:1px solid var(--line);background:var(--card);border-radius:9px;padding:8px 11px;
  font-family:inherit;font-size:16px;color:var(--ink)}
.input:focus{outline:none;border-color:var(--clay)}
.select{border:1px solid var(--line);background:var(--card);border-radius:9px;padding:8px 11px;
  font-family:inherit;font-size:16px;color:var(--ink);cursor:pointer}
.select:focus{outline:none;border-color:var(--clay)}

/* buttons */
.btn{display:inline-flex;align-items:center;gap:6px;border-radius:8px;padding:7px 12px;font-family:inherit;
  font-size:15.5px;font-weight:500;cursor:pointer;border:1px solid var(--line-2);background:var(--card);color:var(--ink);
  transition:.12s}
.btn:hover{background:var(--card-2);border-color:var(--ink-faint)}
.btn:disabled{opacity:.5;cursor:not-allowed}
.btn-primary{background:var(--clay);color:#fff;border-color:var(--clay)}
.btn-primary:hover{background:#403a32;border-color:#403a32}
.btn-ghost{background:transparent;border-color:transparent}
.btn-ghost:hover{background:var(--sand);border-color:transparent}
.btn-danger{background:transparent;color:#a33a2c;border-color:#e6cdc7}
.btn-danger:hover{background:#f7ebe9;border-color:#a33a2c}

/* schedule */
.schedule{padding:6px 30px 50px}
.room{margin-bottom:18px}
.room-head{width:100%;display:flex;align-items:center;gap:10px;background:transparent;border:none;cursor:pointer;
  padding:10px 4px;border-bottom:1px solid var(--line);color:var(--ink);text-align:left}
.room-name{font-size:25.4px;font-weight:600}
.room-count{font-size:13.5px;color:var(--ink-soft);background:var(--sand);border:1px solid var(--line);
  padding:2px 9px;border-radius:20px}
.room-sub{margin-left:auto;font-size:16.5px;font-weight:500}
.room-sub em{color:var(--ink-soft);font-weight:400;font-size:14.5px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:12px;padding-top:12px}

/* card */
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;transition:.12s}
.card:hover{border-color:var(--line-2);box-shadow:0 2px 10px rgba(43,39,34,.05)}
.card-main{display:flex;gap:12px;padding:12px}
.thumb{width:84px;height:84px;border-radius:9px;overflow:hidden;flex:none;background:var(--sand);border:1px solid var(--line)}
.thumb img{width:100%;height:100%;object-fit:cover}
.thumb-ph{width:100%;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;
  gap:3px;color:var(--ink-faint)}
.thumb-ph span{font-size:12px;letter-spacing:.05em;text-transform:uppercase;text-align:center;padding:0 4px}
.thumb-ph.big{font-size:13.5px}
.card-body{flex:1;min-width:0;display:flex;flex-direction:column;gap:5px}
.card-top{display:flex;justify-content:space-between;align-items:center;gap:8px}
.retailer{font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-soft);font-weight:600}
.status-pill{border:none;border-radius:7px;padding:3px 10px;font-size:13.5px;font-weight:600;cursor:pointer;
  font-family:inherit;max-width:150px}
.name{font-size:16.5px;font-weight:600;line-height:1.3}
.meta{display:flex;flex-wrap:wrap;gap:8px;font-size:14px;color:var(--ink-soft)}
.meta .sku{font-variant-numeric:tabular-nums}
.card-foot{display:flex;justify-content:space-between;align-items:flex-end;gap:8px;margin-top:2px}
.price-block{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.qty{font-size:14.5px;color:var(--ink-soft)}
.price{font-size:18px;font-weight:600;font-variant-numeric:tabular-nums}
.line-tot{font-size:14px;color:var(--ink-soft)}
.tag{display:inline-flex;align-items:center;gap:2px;font-size:12.5px;font-weight:600;padding:1px 6px;border-radius:6px}
.tag-edit{background:#ECE6DC;color:#6B5E4B}
.tag-up{background:#F2DAD6;color:#A33A2C}
.tag-down{background:#D9E7D9;color:#356037}
.avail-block{display:flex;flex-direction:column;align-items:flex-end;gap:2px;text-align:right}
.avail{font-size:13.5px;color:var(--sage);font-weight:600}
.avail.oos{color:#A33A2C;display:inline-flex;align-items:center;gap:3px}
.checked{font-size:12.5px;color:var(--ink-soft)}
.approval-row{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:8px;
  padding-top:8px;border-top:1px dashed var(--line)}
.appr-pill{font-size:13.5px;font-weight:600;padding:3px 10px;border-radius:20px}
.approve-btns{display:flex;gap:6px}
.mini{display:inline-flex;align-items:center;gap:3px;border:1px solid var(--line);background:var(--card);
  border-radius:7px;padding:4px 9px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit}
.mini.ok{color:#356037}.mini.ok:hover{background:#e9f1e7;border-color:#356037}
.mini.bad{color:#a33a2c}.mini.bad:hover{background:#f7ebe9;border-color:#a33a2c}
.row-actions{display:flex;gap:4px}
.icon-btn{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:8px;
  border:1px solid var(--line);background:var(--card);color:var(--ink-soft);cursor:pointer;text-decoration:none}
.icon-btn:hover{border-color:var(--clay);color:var(--clay)}
.icon-btn:disabled{opacity:.5}

/* editor */
.editor{border-top:1px solid var(--line);background:var(--paper);padding:14px;display:grid;
  grid-template-columns:1fr 1fr;gap:10px}
.field{display:flex;flex-direction:column;gap:4px}
.field>span{font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);font-weight:600}
.field-wide{grid-column:1/-1}
.editor-foot{grid-column:1/-1;display:flex;justify-content:flex-end}

/* draft */
.draft{margin:0 26px 16px;background:var(--card);border:1px solid var(--clay);border-radius:14px;overflow:hidden;
  box-shadow:0 8px 30px rgba(35,32,27,.10)}
.draft-head{display:flex;justify-content:space-between;align-items:center;padding:13px 16px;border-bottom:1px solid var(--line)}
.draft-head strong{font-size:19.5px}
.draft-status{display:flex;align-items:center;gap:7px;padding:9px 16px;font-size:15px;font-weight:500;background:var(--paper)}
.draft-status.ok{color:#356037;background:#eef4ec}
.draft-status.warn{color:#9A6B16;background:#fbf2e0}
.draft-grid{display:flex;gap:16px;padding:16px}
.draft-thumb{width:120px;flex:none}
.draft-thumb img{width:120px;height:120px;object-fit:cover;border-radius:10px;border:1px solid var(--line)}
.draft-fields{flex:1;display:grid;grid-template-columns:1fr 1fr;gap:10px}
.draft-foot{display:flex;justify-content:flex-end;gap:8px;padding:0 16px 16px}

/* misc */
.empty{display:flex;flex-direction:column;align-items:center;gap:10px;padding:60px 20px;color:var(--ink-soft)}
.foot{text-align:center;font-size:14px;color:var(--ink-soft);padding:20px;line-height:1.6}

/* segmented toggle */
.seg{display:inline-flex;border:1px solid var(--clay);border-radius:10px;overflow:hidden;background:var(--card)}
.seg button{display:inline-flex;align-items:center;gap:6px;border:none;background:transparent;padding:9px 15px;
  font-family:inherit;font-size:15.5px;font-weight:600;color:var(--clay);cursor:pointer;transition:.15s}
.seg button:hover{background:var(--clay-soft)}
.seg button.on{background:var(--clay);color:#fff}

/* spreadsheet table */
.table-wrap{overflow-x:auto;padding:8px 26px 8px;margin-top:2px}
table.sheet{border-collapse:separate;border-spacing:0;width:100%;min-width:1340px}
.sheet thead th{position:sticky;top:0;background:var(--paper);text-align:left;font-size:12px;letter-spacing:.07em;
  text-transform:uppercase;color:var(--ink-soft);font-weight:700;padding:9px 10px;border-bottom:1px solid var(--line);
  white-space:nowrap;z-index:3}
.sheet thead th.r{text-align:right}
.sheet td{padding:6px 10px;border-bottom:1px solid var(--line);vertical-align:middle;background:var(--card)}
.sheet td.r{text-align:right}
.sheet tr:hover td{background:#fffdf6}
.sheet tr.grp td{background:var(--paper);padding:9px 10px;border-top:1px solid var(--line)}
.sheet tr.grp .font-display{font-size:18px;font-weight:600}
.sheet tr.grp em{color:var(--ink-soft);font-size:13.5px;margin-left:10px;font-style:normal}
.sheet tr.subtotal td{background:#fbf6ee;font-weight:700;font-size:14px;border-bottom:2px solid var(--line);
  color:var(--ink);padding:6px 8px}
.sheet tr.subtotal td.r:first-child{color:var(--ink-soft);text-transform:uppercase;letter-spacing:.06em;font-size:12.5px}
.thumb-col{width:46px}
.thumb-sm{width:38px;height:38px;border-radius:7px;overflow:hidden;border:1px solid var(--line);background:var(--paper)}
.thumb-sm img{width:100%;height:100%;object-fit:cover}
.thumb-sm-ph{width:100%;height:100%;display:flex;align-items:center;justify-content:center;color:var(--ink-soft)}
.cell{border:1px solid transparent;background:transparent;border-radius:6px;padding:5px 7px;font-family:inherit;
  font-size:15px;color:var(--ink);width:100%;min-width:54px}
.cell:hover{background:var(--paper)}
.cell:focus{outline:none;background:#fff;border-color:var(--clay);box-shadow:0 0 0 2px rgba(176,91,51,.12)}
.cell.num{text-align:right;font-variant-numeric:tabular-nums;min-width:64px}
.cell.qty{min-width:44px;width:48px}
select.cell{cursor:pointer;-webkit-appearance:none;appearance:none}
select.cell.status{font-weight:600;border-radius:20px;padding:4px 12px;font-size:13.5px;min-width:165px;max-width:none;text-overflow:clip}
.num-ro{font-variant-numeric:tabular-nums;white-space:nowrap;padding:5px 8px}
.num-ro.tax{color:var(--ink-soft);font-size:14px}
.name-col{min-width:240px}.name-col .cell{font-weight:600}
.room-col{min-width:118px}
.type-col{min-width:130px}
.sku-sub{font-size:12.5px;color:var(--ink-soft);padding:0 7px;font-variant-numeric:tabular-nums}
.finish-col{min-width:150px}.dim-sub{font-size:12.5px;color:var(--ink-soft);padding:0 7px}
.src-col{min-width:138px}.src-col .cell{color:var(--clay);font-weight:600;font-size:13px;letter-spacing:.03em;text-transform:uppercase;white-space:nowrap}
.lead-col{min-width:86px}.last-col{font-size:13.5px;color:var(--ink-soft);white-space:nowrap}
.notes-col{min-width:150px}
.price-cell{display:flex;align-items:center;gap:3px;justify-content:flex-end}
.revert{border:none;background:transparent;color:var(--clay);cursor:pointer;padding:2px;display:inline-flex}
.chg{display:inline-flex;align-items:center;gap:1px;font-size:12px;font-weight:700;justify-content:flex-end;padding-right:7px}
.chg.up{color:#A33A2C}.chg.down{color:#356037}
.appr-pill.sm{font-size:12.5px;padding:2px 8px;white-space:nowrap}
.avail.sm{font-size:13.5px;white-space:nowrap}.avail.oos.sm{font-size:13px}
.act-col{width:1%;white-space:nowrap}
.icon-btn.ok{color:#356037}.icon-btn.ok:hover{border-color:#356037;color:#356037}
.icon-btn.bad{color:#a33a2c}.icon-btn.bad:hover{border-color:#a33a2c;color:#a33a2c}

/* discounts */
.savings{color:#356037;font-weight:600}
.tag-disc{background:#E4EBDD;color:#4F6B3A}
.disc-col .net-sub{font-size:12.5px;color:#356037;font-weight:600;padding-right:7px;text-align:right}
.cell.disc{min-width:52px;width:56px}
.disc-panel{margin:12px 0 0;background:var(--card);border:1px solid var(--clay);border-radius:12px;overflow:hidden}
.disc-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:13px 16px;border-bottom:1px solid var(--line)}
.disc-head strong{font-size:18px}
.disc-hint{font-size:14px;color:var(--ink-soft);margin-top:3px;max-width:560px;line-height:1.5}
.disc-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px 16px;padding:14px 16px}
.disc-row{display:flex;align-items:center;justify-content:space-between;gap:8px}
.disc-del{flex:none;color:var(--ink-faint)}
.disc-del:hover{color:var(--clay)}
.disc-add{border-top:1px solid var(--line);padding:14px 16px;background:var(--card-2)}
.disc-add-label{font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-bottom:8px}
.disc-add-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.disc-add-row .input:first-child{flex:1;min-width:180px}
.disc-name{font-size:15px;font-weight:600}
.disc-input-wrap{display:flex;align-items:center;gap:4px}
.disc-pct{width:70px;text-align:right}
.disc-sign{color:var(--ink-soft);font-size:14.5px}

/* dashboard */
.back-link{display:inline-flex;align-items:center;gap:3px;background:transparent;border:none;color:var(--clay);
  font-family:inherit;font-size:14.5px;font-weight:600;cursor:pointer;padding:0 0 6px;letter-spacing:.02em}
.back-link:hover{text-decoration:underline}
.newclient{margin-top:16px;background:var(--card);border:1px solid var(--clay);border-radius:12px;padding:16px}
.newclient-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.newclient-foot{display:flex;justify-content:flex-end;gap:8px;margin-top:12px}
.dash{padding:22px 26px}
.client-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px}
.client-card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px;cursor:pointer;
  transition:.16s;display:flex;flex-direction:column;gap:6px;position:relative}
.client-card:hover{box-shadow:0 10px 30px rgba(35,32,27,.10);border-color:var(--clay);transform:translateY(-2px)}
.client-cover{margin:-18px -18px 14px;height:140px;border-radius:15px 15px 0 0;overflow:hidden;background:var(--sand)}
.client-cover img{width:100%;height:100%;object-fit:cover;display:block}
.client-card.has-cover .client-avatar{display:none}
.ov-cover{position:relative;height:220px;border-radius:16px;overflow:hidden;margin-bottom:18px;background:var(--sand)}
.ov-cover img{width:100%;height:100%;object-fit:cover;display:block}
.ov-cover-name{position:absolute;left:20px;bottom:14px;color:#fff;font-size:30px;text-shadow:0 2px 14px rgba(0,0,0,.5)}
.ov-cover-edit{margin-bottom:18px;max-width:420px}
.ov-cover-label{display:block;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-bottom:7px}
.client-card-top{display:flex;justify-content:space-between;align-items:flex-start}
.client-avatar{width:44px;height:44px;border-radius:12px;background:var(--clay-soft);color:var(--clay);
  display:flex;align-items:center;justify-content:center;font-size:26.5px;font-weight:600}
.icon-btn.del{color:var(--ink-soft);width:28px;height:28px}
.icon-btn.del:hover{color:#a33a2c;border-color:#a33a2c}
.client-name{font-size:22.8px;font-weight:600;line-height:1.2;margin-top:2px}
.client-sub{display:flex;align-items:center;gap:5px;font-size:14.5px;color:var(--ink-soft)}
.client-total{font-size:24.4px;font-weight:600;font-family:'Cormorant Garamond',Georgia,serif;margin-top:8px}
.client-total em{font-size:13.5px;color:var(--ink-soft);font-weight:400;font-style:normal;letter-spacing:.06em;text-transform:uppercase}
.client-stats{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}
.client-stats span{font-size:13px;font-weight:600;padding:2px 8px;border-radius:20px;background:var(--paper);
  border:1px solid var(--line);color:var(--ink-soft)}
.client-stats .st-await{background:#F6E6CC;color:#9A6B16;border-color:transparent}
.client-stats .st-transit{background:#DEE6EC;color:#3E5A72;border-color:transparent}
.client-stats .st-done{background:#D9E7D9;color:#356037;border-color:transparent}
.client-open{margin-top:10px;font-size:14.5px;font-weight:600;color:var(--clay)}

/* tracking */
.track-strip{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:8px;padding-top:8px;border-top:1px dashed var(--line)}
.track-pill{display:inline-flex;align-items:center;gap:4px;font-size:13px;font-weight:600;background:#DEE6EC;color:#3E5A72;
  padding:2px 9px;border-radius:20px}
.track-eta{font-size:13px;font-weight:600;color:#356037}
.track-sub{display:flex;align-items:center;gap:3px;font-size:12px;color:#3E5A72;padding:0 7px;white-space:nowrap}
.rcpt-badge{position:absolute;top:-4px;right:-4px;background:var(--clay);color:#fff;font-size:11px;font-weight:700;
  min-width:14px;height:14px;border-radius:8px;display:flex;align-items:center;justify-content:center;padding:0 3px}

/* editor sections + receipts */
.editor-section{grid-column:1/-1;display:flex;align-items:center;gap:8px;font-size:13px;letter-spacing:.08em;
  text-transform:uppercase;color:var(--clay);font-weight:700;border-top:1px solid var(--line);padding-top:12px;margin-top:4px}
.scan-note{font-size:12.5px;font-weight:500;color:var(--ink-soft);text-transform:none;letter-spacing:0}
.mini.upload{cursor:pointer}
.receipt-list{display:flex;flex-direction:column;gap:6px}
.receipt{display:flex;align-items:center;gap:10px;background:var(--card);border:1px solid var(--line);border-radius:9px;padding:7px 9px}
.receipt img{width:38px;height:38px;object-fit:cover;border-radius:6px;border:1px solid var(--line)}
.receipt-ph{width:38px;height:38px;border-radius:6px;background:var(--paper);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;color:var(--ink-soft)}
.receipt-info{flex:1;display:flex;flex-direction:column;min-width:0}
.receipt-info a,.receipt-info span{font-size:15px;font-weight:600;color:var(--ink);text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.receipt-info a:hover{color:var(--clay);text-decoration:underline}
.receipt-meta{font-size:13px;color:var(--ink-soft);font-weight:400}

/* section tabs */
.tabs{display:flex;gap:4px;margin-top:18px;border-bottom:1px solid var(--line)}
.tab{display:inline-flex;align-items:center;gap:7px;background:transparent;border:none;border-bottom:2px solid transparent;
  padding:10px 14px;margin-bottom:-1px;font-family:inherit;font-size:16px;font-weight:600;color:var(--ink-soft);cursor:pointer}
.tab:hover{color:var(--ink)}
.tab.on{color:var(--clay);border-bottom-color:var(--clay)}

/* contractor totals + table */
.ctr-totals{display:flex;gap:12px;flex-wrap:wrap;padding:18px 26px 4px}
.ctr-tot{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:13px 20px;min-width:170px}
.ctr-tot span{display:block;font-size:13px;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-soft)}
.ctr-tot strong{font-size:27.3px;font-weight:600}
.link-mini{display:inline-flex;align-items:center;gap:2px;background:none;border:none;color:var(--clay);
  font-family:inherit;font-size:13px;font-weight:600;cursor:pointer;padding:2px 0 0}
.link-mini:hover{text-decoration:underline}
.ctr-detail td{background:var(--paper);padding:14px 16px}
.ctr-detail-grid{display:grid;grid-template-columns:1.3fr 1fr;gap:20px}
.ctr-sub-head{display:flex;align-items:center;gap:8px;font-size:13px;letter-spacing:.08em;text-transform:uppercase;
  color:var(--clay);font-weight:700;margin-bottom:8px}
.ctr-sub-head .mini{margin-left:auto}
.pay-row{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.pay-sum{font-size:14.5px;color:var(--ink-soft);margin-top:6px}
.ctr-contact{display:flex;gap:6px;margin-top:8px}
@media(max-width:640px){.ctr-detail-grid{grid-template-columns:1fr}}

/* tracking cell in list */
.track-col{min-width:150px}
.track-cell{display:flex;align-items:center;gap:3px}
.track-cell .cell{min-width:74px}
.track-sub.ship{color:#356037;font-weight:600}
.ship-note{display:flex;align-items:center;gap:6px;font-size:14.5px;color:#356037;font-weight:500;background:#eef4ec;
  padding:7px 10px;border-radius:8px}

/* area manager */
.area-add{display:flex;gap:8px;padding:14px 16px 8px}
.area-list{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px 12px;padding:6px 16px 16px}
.area-row{display:flex;align-items:center;gap:6px}
.area-name{flex:1;font-weight:600;font-size:15px}
.area-count{font-size:13px;color:var(--ink-soft);white-space:nowrap}
.area-empty{padding:18px;color:var(--ink-soft);font-size:15px;font-style:italic}
.area-empty-row td{background:var(--card);color:var(--ink-soft);font-size:14.5px;font-style:italic;padding:10px 12px}

/* ===== project shell: sidebar + content ===== */
.proj-shell{display:flex;min-height:100vh;align-items:stretch}
.proj-nav{width:236px;flex:none;background:var(--sand);border-right:1px solid var(--line);padding:18px 14px;
  display:flex;flex-direction:column;gap:4px;position:sticky;top:0;height:100vh;overflow-y:auto}
.proj-content{flex:1;min-width:0;background:var(--paper)}
.back-link{display:inline-flex;align-items:center;gap:4px;background:transparent;border:none;color:var(--ink-soft);
  font-family:inherit;font-size:14.5px;font-weight:500;cursor:pointer;padding:2px 6px 8px;align-self:flex-start}
.back-link:hover{color:var(--ink)}
.proj-id{display:flex;align-items:center;gap:10px;padding:6px 6px 14px;margin-bottom:6px;border-bottom:1px solid var(--line)}
.proj-avatar{width:36px;height:36px;border-radius:9px;background:var(--clay-soft);color:var(--clay);display:flex;
  align-items:center;justify-content:center;font-size:21.5px;font-weight:600;flex:none}
.proj-id-text{min-width:0}
.proj-name{font-size:16.5px;font-weight:600;line-height:1.2;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.proj-client{font-size:14px;color:var(--ink-soft);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.nav{display:flex;flex-direction:column;gap:2px;margin-top:2px}
.nav-item{display:flex;align-items:center;gap:10px;width:100%;text-align:left;background:transparent;border:none;
  border-radius:8px;padding:8px 10px;font-family:inherit;font-size:16px;font-weight:500;color:var(--ink-soft);cursor:pointer;transition:.12s}
.nav-item:hover{background:var(--card);color:var(--ink)}
.nav-item.on{background:var(--clay);color:#fff;box-shadow:0 1px 2px rgba(43,39,34,.08)}
.nav-item svg{flex:none;opacity:.85}
.nav-foot{margin-top:auto;padding-top:10px;border-top:1px solid var(--line);display:flex;flex-direction:column;gap:2px}
.foot{text-align:center;font-size:13.5px;color:var(--ink-faint);padding:24px;border-top:1px solid var(--line);margin-top:30px}

/* overview */
.ov{padding:8px 30px 50px}
.ov-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}
.ov-stat{text-align:left;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;
  font-family:inherit;transition:.12s}
.ov-stat:hover{border-color:var(--line-2)}
.ov-stat-val{font-size:30.7px;font-weight:500;letter-spacing:-.01em}
.ov-stat-label{font-size:13.5px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-top:6px}
.ov-stat-sub{font-size:14px;color:var(--ink-faint);margin-top:3px}
.ov-cols{display:grid;grid-template-columns:1.5fr 1fr;gap:14px;margin-top:14px}
.ov-card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.ov-card-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
.ov-card-head .font-display{font-size:19.5px;font-weight:500}
.ov-empty{color:var(--ink-soft);font-size:15.5px;padding:14px 0}
.ov-pins{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.ov-pin{aspect-ratio:1;border-radius:9px;overflow:hidden;background:var(--sand);border:1px solid var(--line)}
.ov-pin img{width:100%;height:100%;object-fit:cover}
.ov-detail{display:flex;flex-direction:column;gap:8px;font-size:15.5px;color:var(--ink-soft)}
.ov-detail div{display:flex;align-items:center;gap:7px}
.ov-links{display:flex;gap:8px;margin-top:14px}
@media(max-width:760px){.ov-cols{grid-template-columns:1fr}}

/* pinboard masonry */
.pin-wrap{padding-bottom:50px}
.masonry{column-count:4;column-gap:12px;padding:18px 30px}
@media(max-width:1100px){.masonry{column-count:3}}
@media(max-width:760px){.masonry{column-count:2}}
.pin{position:relative;break-inside:avoid;margin:0 0 12px;border-radius:12px;overflow:hidden;background:var(--card);
  border:1px solid var(--line)}
.pin img{width:100%;display:block}
.pin-ph{height:160px;display:flex;align-items:center;justify-content:center;color:var(--ink-faint)}
.pin-x{position:absolute;top:8px;right:8px;width:26px;height:26px;border-radius:50%;border:none;background:rgba(43,39,34,.6);
  color:#fff;cursor:pointer;display:none;align-items:center;justify-content:center;backdrop-filter:blur(4px)}
.pin:hover .pin-x{display:flex}
.pin-note{width:100%;border:none;border-top:1px solid var(--line);background:var(--card);padding:8px 10px;font-family:inherit;
  font-size:14.5px;color:var(--ink)}
.pin-note:focus{outline:none;background:var(--card-2)}
.pin-cap{padding:8px 10px;font-size:14.5px;color:var(--ink-soft);border-top:1px solid var(--line)}

/* files */
.files-wrap{padding-bottom:50px}
.files-list{padding:16px 30px;display:flex;flex-direction:column;gap:8px;max-width:860px}
.file-row{display:flex;align-items:center;gap:12px;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:9px 12px}
.file-thumb{width:40px;height:40px;border-radius:7px;object-fit:cover;border:1px solid var(--line);flex:none}
.file-thumb.ph{display:flex;align-items:center;justify-content:center;background:var(--sand);color:var(--ink-soft)}
.file-info{flex:1;min-width:0;display:flex;flex-direction:column}
.file-info a,.file-info span{font-size:15.5px;font-weight:500;color:var(--ink);text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.file-info a:hover{text-decoration:underline}
.file-meta{font-size:13.5px;color:var(--ink-soft);font-weight:400}
.chip{display:inline-flex;align-items:center;gap:4px;border:1px solid var(--line-2);background:var(--card);border-radius:20px;
  padding:4px 10px;font-size:13.5px;font-weight:600;cursor:pointer;color:var(--ink-soft);font-family:inherit}
.chip.on{background:#E4EBDD;color:#4F6B3A;border-color:transparent}
.client-upload{display:flex;align-items:center;gap:14px;background:var(--card-2);border:1px dashed var(--line-2);
  border-radius:12px;padding:16px 18px;margin-bottom:6px}
.client-upload>div{flex:1}
.client-upload strong{font-size:16px}

/* seg + tabs (reskinned) */
.seg{display:inline-flex;border:1px solid var(--line-2);border-radius:8px;overflow:hidden;background:var(--card)}
.seg button{display:inline-flex;align-items:center;gap:6px;border:none;background:transparent;padding:7px 13px;
  font-family:inherit;font-size:15px;font-weight:500;color:var(--ink-soft);cursor:pointer;transition:.12s}
.seg button:hover{background:var(--sand)}
.seg button.on{background:var(--clay);color:#fff}

/* ===== selection + export ===== */
.thumb{position:relative}
.sel-box{position:absolute;top:5px;left:5px;z-index:2;width:24px;height:24px;border:none;border-radius:6px;
  background:rgba(255,255,255,.92);color:var(--ink);display:flex;align-items:center;justify-content:center;cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,.12)}
.sel-box.on{background:var(--ink);color:#fff}
.sel-box.tbl{position:static;width:22px;height:22px;background:transparent;box-shadow:none;color:var(--ink-soft);margin-bottom:3px}
.sel-box.tbl.on{color:var(--clay)}
.card.sel{border-color:var(--ink);box-shadow:0 0 0 1px var(--ink)}
.row-sel td{background:var(--card-2)}
.sel-bar{position:sticky;bottom:0;z-index:30;display:flex;align-items:center;gap:10px;margin:18px 30px 0;padding:12px 16px;
  background:var(--ink);border-radius:12px 12px 0 0;color:#fff;box-shadow:0 -4px 20px rgba(0,0,0,.15)}
.sel-bar .btn{background:transparent;border-color:rgba(255,255,255,.3);color:#fff}
.sel-bar .btn:hover{background:rgba(255,255,255,.12)}
.sel-bar .btn-primary{background:#fff;color:var(--ink);border-color:#fff}
.sel-count{font-weight:600;font-size:15.5px}

.export-modal{position:fixed;inset:0;z-index:200;background:rgba(43,39,34,.45);display:flex;flex-direction:column;backdrop-filter:blur(2px)}
.export-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;background:var(--card);border-bottom:1px solid var(--line);padding:12px 20px}
.export-title{font-size:19.5px;font-weight:500}
.export-actions{display:flex;gap:8px;align-items:center}
.export-scroll{flex:1;overflow:auto;padding:30px;display:flex;justify-content:center}
.print-area{background:#fff;width:100%;max-width:840px;padding:46px 50px;box-shadow:0 10px 40px rgba(0,0,0,.2);color:#222}

/* approval / presentation document */
.doc{font-family:'Cormorant Garamond',Georgia,serif;color:#1c1a17}
.doc-hd{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #1c1a17;padding-bottom:14px;margin-bottom:22px}
.doc-eyebrow{font-family:'Jost',sans-serif;font-size:13.5px;letter-spacing:.18em;text-transform:uppercase;color:#9c6244}
.doc-hd h1{font-size:36px;font-weight:500;margin:4px 0 0}
.doc-sub{font-family:'Jost',sans-serif;font-size:15.5px;color:#6b6357;margin-top:3px}
.doc-date{font-family:'Jost',sans-serif;font-size:14.5px;color:#6b6357}
.doc-room h2{font-size:15.5px;letter-spacing:.14em;text-transform:uppercase;font-weight:600;border-bottom:1px solid #ddd5c7;padding-bottom:6px;margin:18px 0 10px;font-family:'Jost',sans-serif}
.doc-item{display:flex;gap:16px;padding:12px 0;border-bottom:1px solid #efe9dd;page-break-inside:avoid}
.doc-thumb{width:96px;height:96px;flex:none;border:1px solid #e6ded0;border-radius:6px;overflow:hidden;background:#f6f2ea}
.doc-thumb img{width:100%;height:100%;object-fit:cover}
.doc-meta{flex:1;font-family:'Jost',sans-serif}
.doc-name{font-size:18px;font-weight:600}
.doc-line{font-size:15px;color:#5c5446;margin-top:2px}
.doc-link{font-size:13.5px;color:#1155cc;word-break:break-all}
.doc-signoff{width:96px;flex:none;font-family:'Jost',sans-serif;font-size:14.5px;display:flex;flex-direction:column;gap:7px;color:#3c372f}
.doc-footer{margin-top:26px;font-family:'Jost',sans-serif;font-size:15.5px;color:#3c372f}

.pres-doc{font-family:'Jost',sans-serif;color:#1c1a17}
.pres-slide{padding-bottom:24px;margin-bottom:24px;border-bottom:1px solid #eee}
.pres-slide-hd{display:flex;justify-content:space-between;align-items:baseline;border-bottom:1px solid #1c1a17;padding-bottom:8px;margin-bottom:16px}
.pres-slide-hd h2{font-size:28.6px;font-weight:500;font-family:'Cormorant Garamond',Georgia,serif}
.pres-slide-proj{font-size:13.5px;letter-spacing:.16em;text-transform:uppercase;color:#9c6244}
.pres-slide-imgs{display:grid;gap:12px;margin-bottom:14px}
.pres-slide-imgs.n1{grid-template-columns:1fr}
.pres-slide-imgs.n2{grid-template-columns:1fr 1fr}
.pres-slide-imgs.n3{grid-template-columns:repeat(3,1fr)}
.pres-slide-imgs.n4{grid-template-columns:repeat(2,1fr)}
.pres-slide-imgs img{width:100%;border-radius:8px;object-fit:cover;max-height:360px}
.pres-slide-prods{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px;margin-bottom:12px}
.pres-prod{display:flex;gap:10px;border:1px solid #e6ded0;border-radius:8px;padding:8px;page-break-inside:avoid}
.pres-prod img{width:62px;height:62px;object-fit:cover;border-radius:5px;flex:none}
.pres-prod .thumb-ph{width:62px;height:62px;flex:none}
.pres-prod-info{display:flex;flex-direction:column;font-size:14.5px;color:#5c5446;gap:1px;min-width:0}
.pres-prod-info strong{color:#1c1a17;font-size:15.5px}
.pres-prod-info a{color:#1155cc;font-size:13.5px}
.pres-slide-note{font-size:15.5px;color:#3c372f;font-style:italic;border-left:3px solid #e6ded0;padding-left:12px}

/* presentations UI */
.pres-wrap{padding-bottom:60px}
.pres-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px;padding:18px 30px}
.pres-card{display:flex;flex-direction:column;gap:6px;align-items:flex-start;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;text-align:left;font-family:inherit}
.pres-card-ico{width:42px;height:42px;border-radius:9px;background:var(--clay-soft);color:var(--clay);display:flex;align-items:center;justify-content:center;border:none;cursor:pointer}
.pres-card-title{font-size:19.5px;font-weight:500}
.pres-card-meta{font-size:14px;color:var(--ink-soft)}
.pres-card-actions{display:flex;gap:6px;margin-top:8px;align-items:center}
.pres-editbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:16px 30px 6px}
.pres-title-input{flex:1;min-width:180px;font-size:19.5px;font-weight:600;font-family:'Cormorant Garamond',Georgia,serif}
.pres-pages{padding:10px 30px;display:flex;flex-direction:column;gap:14px;max-width:940px}
.pres-page{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
.pres-page-hd{display:flex;align-items:center;gap:10px;margin-bottom:12px}
.pres-page-num{font-size:13.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;flex:none}
.pres-page-hd .input{flex:1}
.pres-page-tools{display:flex;gap:2px}
.pres-imgs{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:8px}
.pres-img{position:relative;width:92px;height:92px;border-radius:8px;overflow:hidden;border:1px solid var(--line);background:var(--sand)}
.pres-img img{width:100%;height:100%;object-fit:cover}
.pres-img .pin-x{display:flex}
.pres-add-img{width:92px;height:92px;border:1px dashed var(--line-2);border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;color:var(--ink-soft);cursor:pointer;font-size:13.5px}
.pres-img-row{display:flex;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.pres-img-row .input{flex:1;min-width:160px}
.pres-pin-strip{display:flex;gap:6px;overflow-x:auto;padding:6px 0 10px}
.pres-pin-thumb{width:62px;height:62px;flex:none;border-radius:7px;overflow:hidden;border:1px solid var(--line);background:var(--sand);cursor:pointer;padding:0;display:flex;align-items:center;justify-content:center;color:var(--ink-soft)}
.pres-pin-thumb img{width:100%;height:100%;object-fit:cover}
.pres-prod-hd{display:flex;justify-content:space-between;align-items:center;margin:10px 0 6px;font-size:14.5px;font-weight:600;color:var(--ink-soft)}
.pres-prod-chips{display:flex;flex-wrap:wrap;gap:6px}
.prod-chip{display:inline-flex;align-items:center;gap:5px;background:var(--sand);border:1px solid var(--line);border-radius:20px;padding:3px 5px 3px 10px;font-size:14px}
.prod-chip button{border:none;background:transparent;cursor:pointer;color:var(--ink-soft);display:flex;padding:0}
.pres-prod-panel{border:1px solid var(--line);border-radius:10px;margin:8px 0;overflow:hidden;background:var(--card)}
.pres-prod-search{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--line);color:var(--ink-soft);background:var(--card-2)}
.pres-prod-search .input{flex:1;border:none;background:transparent;padding:2px 0;font-size:15px}
.pres-prod-search .input:focus{outline:none;box-shadow:none}
.pres-prod-search .icon-btn{flex:none}
.pres-prod-list{max-height:300px;overflow:auto;padding:6px}
.pres-prod-group{margin-bottom:6px}
.pres-prod-room{position:sticky;top:0;background:var(--card);font-size:12px;letter-spacing:.1em;text-transform:uppercase;
  font-weight:600;color:var(--clay);padding:7px 8px 4px;z-index:1}
.pres-prod-room em{color:var(--ink-faint);font-size:12px;margin-left:4px}
.prod-pick{display:flex;align-items:center;gap:9px;width:100%;background:transparent;border:none;border-radius:7px;padding:7px 9px;font-family:inherit;font-size:15px;cursor:pointer;text-align:left;color:var(--ink)}
.prod-pick:hover{background:var(--sand)}
.prod-pick.on{background:var(--clay-soft)}
.prod-pick>svg:first-child{flex:none;color:var(--ink-soft)}
.prod-pick.on>svg:first-child{color:var(--clay)}
.prod-pick-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.prod-pick-type{font-size:13px;color:var(--ink-soft);flex:none}
.prod-pick-pin{color:var(--clay);flex:none}
.pres-notes{width:100%;margin-top:10px;resize:vertical;font-family:inherit}

@media print{
  @page{margin:14mm}
  body *{visibility:hidden!important}
  .print-area,.print-area *{visibility:visible!important}
  .print-area{position:absolute!important;left:0;top:0;width:100%;max-width:none!important;box-shadow:none!important;padding:0!important}
  .export-modal,.export-scroll{position:static!important;overflow:visible!important;max-height:none!important;background:#fff!important;padding:0!important;display:block!important}
  .export-bar{display:none!important}
}

/* ===== login + client access ===== */
.login-app{min-height:100vh;background:var(--paper);display:flex;align-items:center;justify-content:center;padding:24px;position:relative;overflow:hidden}
.login-bg{display:none}
.login-bg{position:absolute;width:min(78vh,640px);max-width:120%;opacity:.05;filter:grayscale(1);pointer-events:none;user-select:none;z-index:0}
.login-card{position:relative;z-index:1;width:100%;max-width:430px;background:var(--card);border:1px solid var(--line);border-radius:18px;padding:38px 34px;box-shadow:0 18px 50px rgba(35,33,29,.10)}
.login-brand{display:flex;flex-direction:column;align-items:center;text-align:center;gap:12px;margin-bottom:26px}
.login-logo{width:96px;height:96px;border-radius:18px;background:#fff;border:1px solid var(--line);display:flex;align-items:center;justify-content:center;overflow:hidden;flex:none;box-shadow:0 6px 18px rgba(35,33,29,.08)}
.login-logo img{width:100%;height:100%;object-fit:contain}
.login-studio{font-size:24.4px;font-weight:600;text-transform:uppercase;letter-spacing:.12em}
.login-tag{font-size:13.5px;color:var(--ink-soft);letter-spacing:.04em}
.login-notice{background:var(--clay-soft);color:#7d4a2e;border-radius:9px;padding:10px 13px;font-size:14px;margin-bottom:16px}
.login-tabs{display:flex;gap:4px;background:var(--sand);border-radius:10px;padding:4px;margin-bottom:18px}
.login-tabs button{flex:1;border:none;background:transparent;padding:9px;border-radius:7px;font-family:inherit;font-size:15px;font-weight:500;color:var(--ink-soft);cursor:pointer}
.login-tabs button.on{background:var(--card);color:var(--ink);box-shadow:0 1px 3px rgba(28,26,23,.08)}
.login-pane{display:flex;flex-direction:column;gap:14px}
.login-copy{font-size:15px;color:var(--ink-soft);line-height:1.55;margin:0}
.login-input{font-size:17px;padding:12px 14px;letter-spacing:.04em;width:100%;margin-bottom:10px}
.login-switch{margin-top:14px;font-size:14px;color:var(--ink-soft);text-align:center}
.login-switch span{color:var(--clay);font-weight:600;cursor:pointer}
.login-switch span:hover{text-decoration:underline}
.login-alt{display:flex;align-items:center;justify-content:center;gap:7px;width:100%;margin-top:16px;padding:11px;
  background:transparent;border:1px solid var(--line-2);border-radius:9px;color:var(--ink-soft);font-family:inherit;
  font-size:14px;font-weight:500;cursor:pointer;transition:.12s}
.login-alt:hover{background:var(--sand);color:var(--ink)}
.login-err{color:#a33a2c;font-size:14px;margin-top:-4px}
.login-btn{justify-content:center;padding:12px;font-size:16px}
.login-foot{margin-top:22px;padding-top:16px;border-top:1px solid var(--line);font-size:12.5px;color:var(--ink-faint);line-height:1.5}

.client-badge{display:inline-flex;align-items:center;gap:5px;align-self:flex-start;background:var(--ink);color:#fff;
  border-radius:20px;padding:3px 11px;font-size:12px;font-weight:600;letter-spacing:.04em;margin-bottom:8px}
.dash-brand{display:flex;align-items:center;gap:14px}
.dash-app{background:var(--paper)}
.dash-hdr{position:relative;overflow:hidden;background:transparent}
.hdr-damask{position:absolute;top:-40px;right:-20px;width:230px;opacity:.05;filter:grayscale(1);pointer-events:none;z-index:0}
.dash-hdr>.hdr-row{position:relative;z-index:1}
.dash-logo{width:54px;height:54px;border-radius:12px;background:#fff;border:1px solid var(--line);overflow:hidden;flex:none;display:flex;align-items:center;justify-content:center}
.dash-logo img{width:100%;height:100%;object-fit:contain}

.access-card{position:relative;margin:auto;width:100%;max-width:460px;background:var(--card);border-radius:16px;padding:24px;box-shadow:0 20px 60px rgba(0,0,0,.25)}
.access-hd{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
.access-title{font-size:20.5px;font-weight:500}
.access-copy{font-size:14.5px;color:var(--ink-soft);line-height:1.55;margin:0 0 16px}
.access-toggle{display:flex;align-items:center;gap:10px;font-size:15.5px;font-weight:500;cursor:pointer;padding:10px 0;border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
.access-toggle input{width:17px;height:17px;accent-color:var(--ink)}
.access-code-row{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;margin:16px 0 12px;flex-wrap:wrap}
.access-label{font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-bottom:4px}
.access-code{font-size:27.6px;font-weight:500;letter-spacing:.06em}
.access-code-btns{display:flex;gap:6px}
.access-instructions{background:var(--sand);border-radius:9px;padding:11px 13px;font-size:14px;color:var(--ink-soft);line-height:1.5;margin-bottom:14px}
.access-instructions strong{color:var(--ink)}
.access-preview{width:100%;justify-content:center}
.access-note{display:flex;align-items:center;gap:6px;margin-top:16px;font-size:12.5px;color:var(--ink-faint);line-height:1.45}

/* ===== notifications ===== */
.notif-wrap{position:relative}
.notif-bell{position:relative;width:38px;height:38px;border-radius:10px;border:1px solid var(--line-2);background:var(--card);
  color:var(--ink);display:flex;align-items:center;justify-content:center;cursor:pointer;transition:.12s}
.notif-bell:hover{background:var(--card-2)}
.notif-dot{position:absolute;top:-5px;right:-5px;min-width:18px;height:18px;padding:0 4px;border-radius:9px;background:var(--clay);
  color:#fff;font-size:11.5px;font-weight:600;display:flex;align-items:center;justify-content:center}
.notif-overlay{position:fixed;inset:0;z-index:90}
.notif-panel{position:absolute;top:46px;right:0;z-index:91;width:340px;max-width:88vw;background:var(--card);
  border:1px solid var(--line);border-radius:13px;box-shadow:0 16px 44px rgba(28,26,23,.16);overflow:hidden}
.notif-hd{padding:13px 16px;border-bottom:1px solid var(--line);font-size:17px;font-weight:600}
.notif-list{max-height:380px;overflow:auto}
.notif-empty{padding:26px 16px;text-align:center;color:var(--ink-soft);font-size:14px}
.notif-item{display:flex;gap:11px;width:100%;text-align:left;background:transparent;border:none;border-bottom:1px solid var(--line);
  padding:12px 16px;cursor:pointer;font-family:inherit}
.notif-item:hover{background:var(--card-2)}
.notif-item.unread{background:#FBF6EE}
.notif-ico{width:28px;height:28px;border-radius:8px;background:var(--sand);color:var(--ink-soft);display:flex;align-items:center;justify-content:center;flex:none}
.notif-item.unread .notif-ico{background:var(--clay-soft);color:var(--clay)}
.notif-body{display:flex;flex-direction:column;gap:2px;min-width:0}
.notif-text{font-size:14px;line-height:1.4;color:var(--ink)}
.notif-when{font-size:12px;color:var(--ink-faint)}
.card-notif{display:inline-flex;align-items:center;gap:4px;background:var(--clay-soft);color:var(--clay);border-radius:20px;
  padding:3px 9px;font-size:12.5px;font-weight:600}

/* ===== comment thread + client actions ===== */
.cmt-link{display:inline-flex;align-items:center;gap:4px;position:relative}
.cmt-badge{font-size:11px;font-weight:700;background:var(--clay);color:#fff;border-radius:9px;padding:0 5px;margin-left:2px}
.thread{margin:8px 12px 12px;border-top:1px solid var(--line);padding-top:10px;display:flex;flex-direction:column;gap:7px}
.thread-empty{font-size:13.5px;color:var(--ink-soft)}
.cmt{display:flex;flex-direction:column;gap:1px;background:var(--card-2);border:1px solid var(--line);border-radius:9px;padding:7px 10px}
.cmt.client{background:var(--clay-soft);border-color:transparent}
.cmt-by{font-size:11.5px;letter-spacing:.06em;text-transform:uppercase;font-weight:600;color:var(--ink-soft)}
.cmt.client .cmt-by{color:var(--clay)}
.cmt-text{font-size:14.5px;line-height:1.4}
.cmt-when{font-size:11.5px;color:var(--ink-faint)}
.thread-input{display:flex;gap:6px;margin-top:2px}
.thread-input .input{flex:1}
.mini.active{outline:2px solid currentColor;outline-offset:1px}

/* ===== image lightbox ===== */
img.zoomable{cursor:zoom-in;transition:opacity .12s}
img.zoomable:hover{opacity:.88}
.lightbox{position:fixed;inset:0;z-index:300;background:rgba(20,18,15,.86);display:flex;align-items:center;justify-content:center;padding:34px;cursor:zoom-out;backdrop-filter:blur(3px)}
.lightbox-img{max-width:92vw;max-height:90vh;border-radius:8px;box-shadow:0 24px 70px rgba(0,0,0,.5);cursor:default;background:#fff}
.lightbox-x{position:absolute;top:18px;right:20px;width:42px;height:42px;border-radius:50%;border:none;background:rgba(255,255,255,.14);
  color:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:.12s}
.lightbox-x:hover{background:rgba(255,255,255,.26)}

/* ===== presentation pins ===== */
.pres-stage{position:relative;display:inline-block;width:100%}
.pres-stage>img{width:100%;border-radius:8px;object-fit:cover;display:block}
.pin-dot-wrap{position:absolute;transform:translate(-50%,-50%);z-index:2}
.pin-dot{width:26px;height:26px;border-radius:50%;border:2.5px solid #fff;background:var(--clay);color:#fff;font-size:13.5px;
  font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 8px rgba(0,0,0,.3);transition:transform .1s}
.pin-dot:hover{transform:scale(1.12)}
.pin-dot.on{background:var(--ink)}
.pin-dot.edit{position:absolute;transform:translate(-50%,-50%)}
.pin-pop{position:absolute;top:50%;left:34px;transform:translateY(-50%);z-index:5;width:230px;background:#fff;border:1px solid var(--line);
  border-radius:11px;box-shadow:0 14px 40px rgba(0,0,0,.22);padding:11px;display:flex;flex-direction:column;gap:7px;cursor:default}
.pin-pop.left{left:auto;right:34px}
.pin-pop.up{top:auto;bottom:50%;transform:translateY(50%)}
.pin-pop-img{width:100%;height:120px;object-fit:cover;border-radius:7px}
.pin-pop-info{display:flex;flex-direction:column;gap:2px;font-size:13.5px;color:var(--ink-soft)}
.pin-pop-info strong{font-size:15px;color:var(--ink);font-family:'Cormorant Garamond',serif;font-weight:600}
.pin-pop-price{color:var(--ink);font-weight:600}
.pin-pop-info a{color:var(--clay);font-weight:600;text-decoration:none}
.pin-pop-x{position:absolute;top:6px;right:6px;width:22px;height:22px;border-radius:50%;border:none;background:rgba(0,0,0,.06);
  color:var(--ink-soft);cursor:pointer;display:flex;align-items:center;justify-content:center}
.pres-prod-no{width:22px;height:22px;border-radius:50%;background:var(--clay);color:#fff;font-size:13px;font-weight:700;
  display:flex;align-items:center;justify-content:center;flex:none;align-self:flex-start}

/* presentation editor staging + pinning */
.pin-place-bar{display:flex;align-items:center;gap:6px;background:var(--clay-soft);color:#6f4a2a;border-radius:8px;padding:8px 12px;font-size:14px;margin-bottom:10px}
.pres-drop{border:1.5px dashed transparent;border-radius:12px;transition:.12s;padding:2px}
.pres-drop.drag{border-color:var(--clay);background:var(--clay-soft)}
.pres-edit-imgs{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px}
.pres-edit-stage{position:relative;width:200px;border-radius:9px;overflow:hidden;border:1px solid var(--line);background:var(--sand)}
.pres-edit-stage.placing{cursor:crosshair;outline:2px solid var(--clay);outline-offset:1px}
.pres-edit-stage>img{width:100%;height:150px;object-fit:cover;display:block}
.pres-edit-stage .pin-x{position:absolute;top:5px;right:5px}
.pres-add-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.pres-drop-hint{font-size:13.5px;color:var(--ink-faint)}
.pres-pinsel{display:flex;align-items:center;gap:8px;margin:6px 0 2px;font-size:14px;color:var(--ink-soft);flex-wrap:wrap}
.pres-pinsel .select{min-width:220px}

/* image input */
.imgin{border:1.5px dashed var(--line-2);border-radius:11px;padding:10px;background:var(--card-2);transition:.12s;display:flex;flex-direction:column;gap:8px}
.imgin.drag{border-color:var(--clay);background:var(--clay-soft)}
.imgin-empty{display:flex;flex-direction:column;align-items:center;gap:4px;padding:18px 8px;color:var(--ink-soft);text-align:center;font-size:14px}
.imgin-browse{color:var(--clay);font-weight:600;cursor:pointer;text-decoration:underline}
.imgin-has{position:relative}
.imgin-has img{width:100%;border-radius:8px;display:block;max-height:220px;object-fit:contain;background:#fff}
.imgin-clear{position:absolute;top:6px;right:6px;width:24px;height:24px;border-radius:50%;border:none;background:rgba(0,0,0,.55);color:#fff;cursor:pointer;display:flex;align-items:center;justify-content:center}
.imgin-url .input{width:100%;font-size:14px}
.draft-photo{width:230px;flex:none}
.draft-photo-label{font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-bottom:6px}
.draft-photo-label .link-mini{text-transform:none;letter-spacing:0}

/* ===== finance ===== */
.fin-wrap{display:flex;flex-direction:column;gap:16px}
.fin-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.fin-tile{background:var(--card);border:1px solid var(--line);border-radius:13px;padding:15px 16px}
.fin-tile.good{background:var(--clay-soft);border-color:transparent}
.fin-tile-label{font-size:12px;letter-spacing:.07em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;margin-bottom:7px}
.fin-tile-val{font-family:'Cormorant Garamond',serif;font-weight:600;font-size:27px;line-height:1;color:var(--ink)}
.fin-tile-sub{font-size:12px;color:var(--ink-faint);margin-top:5px}
.fin-form{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;background:var(--card);border:1px solid var(--line);border-radius:13px;padding:16px}
.fin-form .field{display:flex;flex-direction:column;gap:4px}
.fin-form .field>span{font-size:12px;letter-spacing:.05em;text-transform:uppercase;color:var(--ink-soft);font-weight:600}
.fin-form .field-wide{grid-column:1/-1}
.fin-form-foot{grid-column:1/-1;display:flex;justify-content:flex-end;gap:8px}
.fin-table th{font-size:11.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-soft);font-weight:600;padding:9px 12px;text-align:left;border-bottom:1px solid var(--line)}
.fin-table th.r,.fin-table td.r{text-align:right}
.fin-table td{padding:11px 12px;border-bottom:1px solid var(--line);font-size:14.5px;vertical-align:middle}
.fin-row{cursor:pointer}
.fin-row:hover{background:var(--card-2)}
.fin-notes{max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pay-status{display:inline-block;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600}
.track-pill.pay{background:var(--clay-soft);color:var(--clay);border-color:transparent}

@media(max-width:560px){
  .draft-grid{flex-direction:column}.draft-fields{grid-template-columns:1fr}
  .editor{grid-template-columns:1fr}.cards{grid-template-columns:1fr}
}
`}</style>
  );
}
