<?php
session_start();

// Restrict access to only admins
if (!isset($_SESSION['role']) || $_SESSION['role'] !== 'admin') {
    header("Location: login.php");
    exit();
}
?>
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Dashboard</title>
    <link rel="stylesheet" href="/itonatocuescano/CSS/fonts.css">
        <link rel="icon" type="image/png" href="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxIQEBITEBEQEBEQDxUQEBAQEREVFRUXFRIWFhcRFRMYHSggGBolHxUVITEhJSkrLi4uFx8zODMtNygtLisBCgoKDg0OFQ8PGjcdFRk3KzctKystKystKysrLSsrLS0tLSsrKystNysrKysrLS0tLSs3KysrLSstLSsrKystK//AABEIAMkA7AMBIgACEQEDEQH/xAAcAAEAAQUBAQAAAAAAAAAAAAAABgEEBQcIAwL/xABREAABAwECBwwGBQkHAwUBAAACAAEDBAURBhIWF1OT0QcTITFBUVJUYZGS0iIycYGU0xRVZILBCBUjQmJyoaKxGCRDY6TC4YSy4kSjs/DxM//EABgBAQADAQAAAAAAAAAAAAAAAAABAgME/8QAKBEBAAECBAcBAAIDAAAAAAAAAAECEQMiUWESEyEjMUFSFEJiMoLw/9oADAMBAAIRAxEAPwDeKIiAiIgIiICIiAiIgIigmFO6nZ1A7jvn0mZv8KnuK5/2z9Ue937EE7VHdc427u418zu1NHDSC/E7Nv0jfeP0P5VBLWwlrKu/6TVTzM734hyk4e4L8VvcyDrKtwooYHxZa2kjJuMTqImLw416wtXuoWPG9xV8bv8A5YSyN4gB2XKF6og6nfdesbrj/D1XkX0G61Yz/wDrbvbBVN/sXKyIOvKXD+y5fVtClb9+UY3/AJ7lnaKtjmFjhkjlB+IojEx8QriZXVDXSwHjwyyRG3EcRkBN94UHbCLmfBzdktGmdmnIa2JuBxmbFku7JR/qeMtt4P7pFNace90k0VLWu3ow1gETO/MGKY4/ue/sQT1Fqe1MJMJ4JHD83UUzNxSQBMYE3O36Zn72ZWmWOFH1RBqpfnINxotO5Y4UfVEGql+cmWOFH1RBqpfnINxItO5Y4UfVEGql+cmWOFH1RBqpfnINxItO5Y4UfVEGql+cq5YYUfVMGql+cg3CiIgIiICIiAqL4lkYWcidhFmvd3e5mZYOPDCjI2AJSMnJhFhjN73Liu4FE1RHksytRWgBMF98herGPCT9t3I3a/Avi07QjpoTmnMYo4xxjIn4Gb8X7OVXl3/6uaN2LDp7QqHp4D/udOdwuz8EsjcDy9o8jd/KpDdF3Uqi0COGncqej4RxWe6SVulITcn7De+9a2vVb1RAREQEREBERAREQFVnVEQbc3PN2CWmcYLRI56fgEZ+OWNucuWQW7/bxLbGFeFx0lKNXTUzV1M4Y5yxTsziPTuxHvDt5OZclrYm5buhHZsm8zu8lDK9xg/pb078coNzc7cvtQTJ/wAoIPq4/ih+Wn9oIPq4/im+WprPgRYbR7+VNSDETMTSueJG7HxOxY11z3q0yZwb6Fm/Ej50EV/tBB9XH8U3y0/tBB9XH8U3y1Kcl8G+hZ3xI+dHwYwb6Fm/Ej50EWb8oAPq4/iW+Wtw2XOUkERmG9HJEJnE74zgRCzuDly3X3e5Q+zsD7AlO6CCilMWx8WObfHZmdvSxWN+DhZTtAREQEREFF5yzCDXkQi3OTszL1UFt7Besq5cc54WFndgjZzuFuzg9bndUrqmmOkXTEXSav8Aos44kpxmN992+3Nf7nXjZthUcZNJBEDE1+KYkRdj8rqG5uZ9LB/PsU9sahangjib/DG535343fvvVKJqqnNTZM2jxKH7s2Ez0FmkMZYs9W+8ROz8LD/iSe5uC/kcxXLrrYe7dhD9MtQ4xe+Kjb6ODcjnf+kfv4PuLXa2VERZCw7Nkq6iKCJsaSaRowbkvflfsbjfsvQfVj2NUVcjR00Ek5vw4oDfc3SIuIW7X4FOqPcTtSRrzamgv/VlmvdtWJMt7YG4MQWZTBBAzX3M8sjtcUp8pl+DcjJhZWVMcYjSREchO95iONiNtdRM2i5EXaTbcHtHT0Osn+WmYe0tPQayf5Kmb09rPx/TOH9s2/FGgtb7b4zXP+ifmWnL3QzMNaPWKHxz/KTMPaOnoNZP8lTTerW+2+I03q1ft3iNP0T8ycvdC8w9o6eh1k/y0zD2lp6DWT/JU03q1vtviNN6tb7b4jT9E/MnL3QvMNaPWKHWT/LTMPaOnoNZP8lTTerW+2+I03q1ft3iNP0T8ycvdC8w9o6eh1k/y1XMPaPWKHxz/LUy3q1vt3iNV3q1vtviNP0f1k5e7M4CYI1dPQzUNplTVNMTYsQxlK5CJevG7kI8F9zs7cLO79l0bLcDpr+CtqG5m3sHuV5vVrfbfEao8VrfbfEafo/rKOXusX3Aafr0+qDavuHcDp2f0q2d25WGMBf3O96n+B9XUlE8dVHKJxvwSSC7Y4+3nZSNdFM3i6sxZHcEsDaSy43GlB2I7t8lN8aQ7ukXN2NcykaIpQIiIKK3rqsYYykN7hBr3u4XfsbtXu73KNyYbUTO7b4T3Pde0Zuz+xVqqinzKYiZRWpw6rHJ3AAEL3xRcCJ2bkxi5155c13+Xqv+VLcuqLpnqjR8OaLpnqjXNa/82n+qxwMtyrq5SaXFaIBvJ2DFvcuIf6usxhrbrWfQVFQ92NHG+9s/LIXogPe7LLU07SAJjfimLEN7XPc/FwLSP5RWEN5U9CD8A/3mf2veMY92O/vZdNETEWmbs56tLSSORORO7kTuRO78LvxrzRFZAt1fk74OY0k1cbXtF/d6d36RNfIXuZxb77rTAC7vczO7u9zM3G/Yuv8AAew2oLPpqfiKONnlu5ZD9OT+Lv7kGed1Cpt0OFndhhkJmd2YmIGZ26XCpfVYji4yOzCTOL3vdwOsHk1Z2jh1heZZYnF/GbLU29sZnGh0E3eG1UzjQ6CXxAspk3Z2jh1peZVybs7Rw6wvMs7Yv1C2TRis40Ogl8QKucaHQS94LKZNWdo4tYXmTJqztHFrC8yd36gyaMXnGh0E3fHtTONDoJe8FlMmrO0cWsLzJk1Z2ji1heZO79QZNGLzjQ6CXvBM40Ogl749qymTVnaOLWF5kyas7RxawvMnd+oMmjGZxYdBL3htVM4sOhl7wWUyas7RxawvMmTVnaOLWF5kti/UGTRi840Ogl749qZxYdDL4gWUyas7RxawvMmTVnaOLWF5kti/UGTRi23RYdBLdz4wKZQzMYsQuziQsTO3Kz8SwOTNnaOLWF5lmqKMBAQiuxAbFZhfGubmvWmHxfylWq3pcoiLVURF8k/BwcPBwMgtrQo2niKMiIWNsUnB7iu5r1G831L0p/GHlWHrLCtSSQjeTFxycsUagmEf2WHsZeOTVqaU/iS2rmrqvPWi7SIt4lns31L0p/GHlVR3P6Vn9ad2Z77nMLn/AJVgHwatTSn8SW1STA+yKmB5CqpCInZhAXkc2ZuV/bxKKKaZn/CxMzqkc0gxg5E7CAC5E78DMItwv3LjzC22ir62oqSv/TSuQs/GIcQB7mZmXQe7lhB9FswogK6WtLeBZn4d745i9l3oP++uY11MxERkE73HLB+m2rDjNfFTf3mS/i9C7EHxuHuvXU61duBWB9Hs8qkmukrTxh7IgvEG78d/Y7LZVVMwRmbtfiA5uzcuK16CMYZYOz1phvZxDHGL3MeNe5FxvwC/MKj2bup0lP3yeVeRYeVnNC3ZvZbVXL2s5ofAW1cVdWFVN5u1iKo6Q9c3dTpKfvPyqmbup0lP3yeVeWXtZzQ+AtqZe1nND4C2qtsHdbO9s3dTpKfvPypm7qdJT95+VeWXtZzQ+Atqpl/V80PgLaotg7md7Zu6nSU/eflTN3U6Sn7z8q8svqvmh8BbUy+rOaHwFtS2DuZ3rm7qdJT95+VM3dRpKfvPYvHL6r5ofAW1Mvqvmh8BbU7O5ne2bup0sHeexM3dTpYO89i8svazmh8BbUy8rOaHVltTs7md7Zu6jSwd5+VM3dRpYO8/KvLLys6MPgLamXdZzQ+AtqdnczvTN3UaWDvPyqT4HWBLRNK0hgTSOLjiOXA7Y1/G3sUUy8rOjD4C2qQYGYRz1cphMwMIx47OIu3DePb2rXCnD4svlSrit1TJERdbNRlE8L8Kyo5AjiADJwxyx3K4Wxrm4vYSlbvcsBNb1nE95ywE/K5Bjfgs8TxaJsmPKK5xZ9DD3ntTONPoYe89qk/56szp02r/APFU/PVmdKm1X/isLVfbS8fKNR7oVQRMIwwu7uzCzOfC7+9bFivxWxrsa5r7uK/luWKs2ropjug3gjFsf0QZnZmf1r7u1X9pVgwQyzF6kMRyl7AFyf8Aot8OJiOs3UqmPUObt3K3/pVqFGJXxUY7wLM/Bj+tI/tv9H7i1yrisqClkOU3xjlkKQy5yIsZ373VutFRERB01uYYdUMln00Jzw089PCMJxTSDHe4NisQOVzFfdfwcV62FDUAbXgYkz8okLsuI1Vn93s4EHb9zdircy4japPpl4iVfpMnTPxEg7aubsS5uxcSPUn0yf7zr530uk/e6Dty5uxLm7FxHvpdJ+9030ud+90S7dubsVcVuZu5cRtUG365N9516jXytxSyNdxXGTJZDtbFbmbuTFbmbuXGUFv1YepV1IXdGeVvxWYs/dHtaD1K+oe7kmIZm/8AcZ0sOtCa7k9zXKwqLTGO/HinZm43GEjb+S9aKsfd2rI3ZqmCCoFm4XDGhN/fwj/BltTA3dHorTuCIiiqLr/o8zMxPdx4j8R+7h7FFhenhrRM7s5kztxs8R3/ANEy2odIWqk2LMV9lwztdNEEnaQ8LewuNli8i6HQvrJfMs6or9WWi3t55bUOkLVSbFkbHtuCrxt4Jy3tmxrxIbsbi4/YrHIqh0L6yXzLIWVY0NLjbwDjj3Y15EV93F6z9qimK75rE29MmiItlVtXQb5GYMTi5g4YzNe7YzXXqG5uA6werHasthpLVDEDUjSObm7kUTXuzC3q++/+Ch++2x9r8P8AwufGqpvaabr0xPqWazcB1g9W21M3AdYPVttWG322Ptnh/wCE322Ptfh/4WN6PiVs2qY4NYKhRGZtIUjmGJ6TM1zY16+t0C/81Whdx/QKju3or/4LC4J1Ve091UFQURC44xi9wvxs/F7lNKmAZAIDZiAxcCF+JxJrnZdWFMTT0izOq9+riRFJ8OME57KqjhkZ3jd3eCa70ZA5CxudvRvbkf3O8YuWiBERAREQEREBERAREQEREBERAVxS1BxGJxkQGBMYEL3ELjxEzq3RB17uf2+9oWbT1JXNJIDjKzcDY4E4G/YzuN/vUlWv9wyncLFgd/8AEklka/m31x/2rYCAiIgIiII7hFhSFEYgcZm5BjNiuLNxu3L7Fis40Ogl8QKQ2tZFLOTPUABEzYouRuL3e52VjkzZ2ji1p+ZYV8y+WYsvFrdWLzjRaCXxAmcaLQS+IFlcmbO0cetPzJkzZ2jj1p+ZVti/UJvToxWcaLQS+IFUd0WG/hglbne8HWUyZs7Rx60/MqZM2do4taXmS2L9QZNFpV23ZlfE8VTvZRlxx1AO1z8+NxM/az3qC2vuR2VM7lSV7U178AFJFMDfste7H3k62K2DNnaOLXH50yZs7Rx60/OtImv3ZWbNMVG4jUX/AKGvoJG5zIwf+VjVsW4laLcUtAX7sx/jGt4ZM2do4taXmTJqztHFrT86txSjo0Oe4xajcQ0xfuzj+Ny8n3HLW0ML/wDURbVv7JuzuhHrj86pk3Z3Qj1x+dLydGgczlr6GH4iLamZ21tFD8RFtW/sm7O6EeuPzpk3Z3Qj15+dOKTo0GO49ar/AOHTt7aiL8FVtxy0+VqVvbUAt95NWdo49cfnTJqztHHrj86Xk6NC5mrT56T4gdi+x3F7Tfiekf2T3/7VvfJqztHHrj86rk3Z3Qj1x+dLydGij3E7VZuKmfsabhfvZRvCDAS0KAceqpTCO+55QcJAb2kDvd77l1jQRRAOLEV7Nyb4R3d7vcvG33jakqHmueJqeR5WficMQsb+Cuhxc6KvsVEBXNDSFNLHFG2McsgxA3ORkwsPe6t1svcFsH6TaW/k18dFHvt/Jvh+jGP/AHv9xB0PYlmjS00MAepBEEY9uKN1/vV+iICIiAiIgimGWDR1rxlGQA8bExY7vw33XcTP2qOZu6nSwd57FOsIqmWKmkkgYXkBmK4mvZ2xmv8A4XqBZd1vRh1ZeZc2LThxN6mlE1W6PvN3U6WDvPYmbup0sHeexfGXdb0YtWW1Uy7rejFqy2rDs7rZ3pm7qdJB3nsTN3U6SDvPYvjLut6MWrLamXdb0YtWW1OzuZ33m7qdJB3nsTN3UaSn7z2L4y7rejFqy2pl3W9GLVltTs7md95u6jSU/eexVzd1Gkg7z2Lzy8rOjFqy2pl5WdGLVltTs6SZ33m7qNJT957Ezd1Gkp+89i+cu6zoxastqpl3WdCLVl5k7O5nfebuo0lP3nsTN3UaSn7z2L5y7rOjFqy2pl3WdGLVltTs7md9Zu6jSU/eexM3dRpKfvPYvjLus6EWrLzJl5WdGLVltTs7pzvvN3UaSn7z2Jm7qdJT95+VfGXlZ0YtWW1Vhw2rjIRAYnIyYRbey4XIrm5VMRgz6lGdJsDcGjonkKQgIjZmbEvuZmvv42bsWB3c7e+i2WUQvdJWlvI3cbA3pSP7Lrg++thxMWK2Nc5YrYzs1zO/Kuad3LCD6XaZRAV8VEO8Dc/BvnrSF7b7h+4u2mmKYtDKZv5a4REVkC6f3ELB+iWWEhjdJWE9Qd7cOJ6sY+zFbG++uesD7Fevrqema/8ATSsxu3IA+kZe4WJdhwxCAiIswiIsIs3EzNwMyD1REQEREBERBRYO08JqankeOUiExZnuxCfgLmdZxYS3sG4axxeTHEgZxYgcWd2fke9n/wDrqlfFbL5TFr9VrltRaQtVJsTLai0haqTYrTN7S6SfxB5UzeU2kn8QeVY3xtIXyLzLWi0haqTYmWtFpC1UmxWebym0k/iDyqub2l6c/iDypfG0gyrvLWi0haqTYmWtFpC1UmxWebym0k/iDyqub2m6c/iDypfG0gyLvLai0haqTYmW1FpC1UmxWeb2m0k/iDyqub2l6c/iDypfF0gyrvLWi0haqTYmWtFpC1UmxWmb2l6c/iDypm9penP4g8qXxtIMq7y1otIWqk2JlrRaQtVJsVnm8punP4g8qrm9penP4o/Il8bSDKu8taLSFqpNiZa0WkLVSbFZ5vaXpz+KPyJm9ptJP4g8qXxtIMq7fDeh0haqTYsnZFsQ1TEULuTA7M7uJDwv7Vgc3tNpJ/EHlUgsWyY6SLe4r7sZydye93d+dXo5l81rK1W9LfC22hoKKoqSu/QxOQs/Kb+iA+8nFlx7UTEZEZu5GZORk/GTk97u/vW7PyisILmgoQLjf6TOzP7RjF/53u7BdaMWyoiLIWNZ51U8UETXyTyDGDcN15Pde936vO/IzOg3N+T1gzcMtfI3Cd8FNf0W4ZD973N90luxY6w7Ljo6eKnia4IIxjHtu4yftd+F/asigIiICIiAiIgIiIIda+CtRJKRQVkkYE+MwEcvovyi1z8Sssjq7r5eOfar3C/Biurj/Q2pLQwszM0UEVxO/KRStIxP7OBv6vE6jcmrmEnjtyrI8V8US38Wd+RnJpnub3OspwqZm60VTDP5H1/Xi1kyo+B9f159ZMtBWxaFp0kpw1FTWxSxvikD1E3B/NcTO3CztwOsflNXddrPiZtqcmk4pdG5H1/Xn1s6ZH1/Xn1s65zbCiubirqxv+pm2quVVf1+s+Jm2pyKTil0XkfX9efWzpkfX9efWzrnTKqv6/W/FT+ZMqq/r1b8TN5k5FJxS6MbBCv68+smTJCv68WsmXOmVVf16t+Jm8yZVV/Xq34mbzJyKf8ApOOXReSFf159ZMqZH1/Xn1sy50yqr+vVvxM/mTKqv69W/Ez+ZOTScUujWwOruvFrJ9q9ocDKj9e0Zvuuf4mubMqq/r9b8VP5lQ8KK5+OurHbmepmf/cp5NJxS61sqyBpmveaaV7uEppXdm9g8TKNYZbqFDZ8ZMEgVVRc7DBCYkzP/mG3ADfx7FzBU1kkj/pJJJO0zIv6rwZXiIjwqyNuWrLWVElROWNLMWOb8TNzCzczNczexY1SDBTBGrtOTEpYsa715T9GOPtM/wAGvfsW0g/J/a5sa0XZ7uFmpb2/+RSMBuZUFiwMNTaVXBJM7fo6UxMgj/akbFuM+zibtfi3Tgu1l1P6ez4qQt6Nw32GnAHEsXhFixGfiL+K1x/Z/H6xL4VvmLZ2BmDcdmUcdNG+PiORFI7YrmRFe5O3c3sZkGfREQEREBERAREQFF90Svqaezpyoo5pKkmaKFoIzkMXMrnPFFn4mve/nZlKFrbC3ddpbPq5KZ4JpiiuYzjIGFidsbF9Lla9Bqv854Uc1sfDz+RV/OeFHNbHw8/kU6z+0vU6nxxJn9pOp1PjiQavtqy7brSE6qktKcgHEEjpZ3dmvvxfU51grSsKqpmYqilqYBJ8UXmhONnfotjMt35/aTqdT4otqt6zdwoJwKOagmljJriCR4SEvaLoNCuyKRYU1NnTEx0EFTTXv6cUpgcbfuF6ze+9R1ARLkuQES5LkBEuS5ARGZSLBiezY3xq+GqqHYvRihkCON2/bL1+52QYmz7NmqCxaeGWc2a9xhAzK72CyyWRlpfV1d8JP5VtWyd2OzaOPe6azZIAb9WPehv7X5SftdX+fyk6nU+KLag11Zr4RU0QxU8NqxRBfigFNMzNfx/qK5/OmFHNbHw0/kWwKXdxgmMY4qCrkkN7hAHAid+iItwutm2TVSyxsc0BUxFwtEZgZM37WJwM/Ze6DW25RR2xNI9RadRWRwx3hHTTNiPKTt6xiTX4g38HO93Iz37YREBERAREQEREBERAUKrNyyyZpDklpXOSUyMyeoqryIivcv8A+imqIIJmhsXqT/E1XzEzQ2L1J/iar5inaIIJmhsXqT/E1XzF9Zo7G6l/qKr5inKIIJmhsXqT/E1XzEzQ2L1J/iar5inaIII25FYvU3+JqvmKk25DY5NwUjj2jUVF/wDE1PEQa3LcUsl/1aluxp3/ABZesW4zZA8cMp/vTy/7XZbDRBCG3J7G6k3t36ov/wC9fL7kljdTf4mq+YpyiCC5orF6l/qKr5iZorF6l/qKr5inSIIJmhsXqT/E1XzFXNFYvUv9RVfMU6RBg8H8FKKz2dqOnjhd/WJryMm6LyHeTt2XrOIiAiIgIiICIiAiIgIiIP/Z">
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
    background: linear-gradient(to right,rgb(144, 105, 6),rgb(0, 0, 0)); /* Gradient background */
    padding: 1.5rem;
    height: 100vh;
        }
        aside li a:hover {
            box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.2);
            transform: scale(1.05);
            transition: all 0.3s ease-in-out;
        }
    </style>
</head>
<body class="bg-gray-100">

    <div class="flex">
        <!-- Sidebar -->
        <aside class="w-64 bg-stone-900 h-screen p-6 text-white fixed left-0 top-0">
            <h2 class="text-2xl font-bold mb-6">Admin Dashboard</h2>
            <nav>
                <ul>
                    <li class="py-3"><a href="/itonatocuescano//Admin_dashboard.php"class="transition duration-300 ease-in-out hover:bg-yellow-700 text-xl block rounded">Home</a></li>
                    <li class="py-3"><a href="Company_mission.php"class="transition duration-300 ease-in-out hover:bg-yellow-700 text-xl block rounded">Our Mission</a></li>
                    <li class="py-3"><a href="Company_vision.php"class="transition duration-300 ease-in-out hover:bg-yellow-700 text-xl block rounded">Our Vision</a></li>
                    <li class="py-3"><a href="Company_goal.php"class="transition duration-300 ease-in-out hover:bg-yellow-700 text-xl block rounded">Our Goal</a></li>
                    <li class="py-3"><a href="logout.php" class="transition duration-300 ease-in-out hover:bg-yellow-700 text-xl block rounded">Logout</a></li>
                </ul>
            </nav>
        </aside>

     
        <!-- Main Content -->
        <main class="flex-1 p-8">
            <div class="max-w-4xl mx-auto text-center bg-white text-gray-900 p-10 rounded-lg shadow-lg">
                <h1 class="text-4xl font-bold text-stone-900 mb-6">Our Goal</h1>
                <p class="text-lg leading-relaxed">
                    "At <span class="font-semibold"><a class="transition duration-300 ease-in-out hover:bg-yellow-800 text-xl rounded" href="https://www.facebook.com/systechintegration">Systech Integration & Security Solutions, Inc.</a></span>, we envision a future where 
                    <span class="text-yellow-600 font-semibold">innovation and excellence</span> redefine industry standards. 
                    Our goal is to <span class="text-yellow-600 font-semibold">empower individuals and businesses</span> by 
                    providing cutting-edge solutions, fostering creativity, and driving 
                    <span class="text-yellow-600 font-semibold">sustainable growth</span>. 
                    Through collaboration, integrity, and relentless pursuit of excellence, we aim to be the leading force shaping a brighter, more connected tomorrow."
                </p>
            </div>
        </main>
    </div>
    </div>
</body>
</html>
